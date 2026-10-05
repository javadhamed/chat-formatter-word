#!/usr/bin/env python3
"""Build ChatFormatter.dotm (or .dotm for macOS) from the VBA sources.

Works on any platform with Python 3.8+ and no Word, PowerShell or COM.

Why this exists
---------------
The template is an OOXML zip whose ``word/vbaProject.bin`` is an OLE compound
file.  ``tools/build.ps1`` produced that file by driving Word over COM, which
only works on Windows.  Here we instead patch the VBA module streams of a seed
``vbaProject.bin`` that Word itself produced, so every type library reference
in the ``dir`` stream stays byte identical and the template keeps compiling in
Word on macOS just as well as on Windows.

Only module source changes; ``dir``, ``PROJECT`` and ``PROJECTwm`` are copied
through untouched, which is why the module list and names must stay the same.

Usage
-----
    python3 tools/build_dotm.py                     # -> ChatFormatter.dotm
    python3 tools/build_dotm.py --platform mac      # macOS template name
    python3 tools/build_dotm.py -o dist/mac/ChatFormatter-Mac.dotm
"""

import argparse
import io
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import vbabuild  # noqa: E402

try:
    import olefile
except ImportError:  # pragma: no cover
    sys.exit("This script needs 'olefile'.  pip install olefile")


ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src")
SEED = os.path.join(ROOT, "tools", "seed.dotm")

# Module stream name -> (source file, strip Attribute lines)
MODULES = {
    "ChatFormatter": "ChatFormatter.bas",
    "ThisDocument": "ThisDocument.cls",
}


def read_source(name):
    """Read a .bas/.cls into the exact text Word stores in the module stream.

    Two things matter here and getting either wrong yields a template that
    loads but exposes no macros:

    * the ``Attribute VB_Name`` line must survive.  Word identifies the
      module by that name; dropping it leaves the project without usable
      public entry points.
    * the ``VERSION 1.0 CLASS / BEGIN ... END`` header is export metadata,
      not source, and must be dropped.
    """
    path = os.path.join(SRC, name)
    with open(path, "r", encoding="utf-8") as fh:
        text = fh.read()
    text = text.replace("\r\n", "\n").replace("\r", "\n")

    lines = text.split("\n")
    out = []
    skipping = False
    for ln in lines:
        if re.match(r"^VERSION\s+1\.0\s+CLASS", ln, re.I):
            skipping = True
            continue
        if skipping:
            if re.match(r"^END\s*$", ln, re.I):
                skipping = False
            continue
        out.append(ln)

    # VBA requires CRLF inside the module stream.  With bare LF the whole
    # module is seen as a single line and the project fails to compile, so
    # the template loads as an add-in but exposes no macros.
    return ("\r\n".join(out)).encode("latin-1", "replace")


def build(output):
    if not os.path.exists(SEED):
        sys.exit("Missing seed template: %s" % SEED)

    with zipfile.ZipFile(SEED) as zin:
        names = zin.namelist()
        payload = {n: zin.read(n) for n in names}

    if "word/vbaProject.bin" not in payload:
        sys.exit("Seed has no word/vbaProject.bin")

    ole = olefile.OleFileIO(io.BytesIO(payload["word/vbaProject.bin"]))

    streams = {}
    for entry in ole.listdir(streams=True, storages=False):
        path = "/".join(entry)
        if ole.get_type(path) == olefile.STGTY_STREAM:
            streams[path] = ole.openstream(path).read()

    if "VBA/dir" not in streams:
        sys.exit("Seed has no VBA/dir stream")

    dir_plain = vbabuild.decompress(streams["VBA/dir"])
    offsets = vbabuild.parse_module_offsets(dir_plain)
    if not offsets:
        sys.exit("Could not read module offsets from VBA/dir")
    print("  seed module offsets: %s" % offsets)

    tree = {"PROJECT": streams.get("PROJECT", b""),
            "PROJECTwm": streams.get("PROJECTwm", b"")}
    vba = {"dir": streams["VBA/dir"],
           "_VBA_PROJECT": streams.get("VBA/_VBA_PROJECT", b"")}

    for module, filename in MODULES.items():
        src_path = os.path.join(SRC, filename)
        if not os.path.exists(src_path):
            sys.exit("Missing source: %s" % src_path)
        source = read_source(filename)
        # Keep the existing performance cache so the TextOffset recorded in
        # dir stays valid; only the compressed source that follows changes.
        head = streams.get("VBA/" + module, b"")[:offsets.get(module, 0)]
        vba[module] = head + vbabuild.compress(source)
        print("  %-16s source %6d bytes -> stream %6d bytes"
              % (module, len(source), len(vba[module])))

    tree["VBA"] = vba
    payload["word/vbaProject.bin"] = vbabuild.write_cfb(tree)

    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, payload[n])

    print("Wrote %s (%d bytes)" % (output, os.path.getsize(output)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--output", default=None)
    ap.add_argument("--platform", choices=["windows", "mac"], default="windows")
    args = ap.parse_args()

    if args.output:
        out = args.output
    elif args.platform == "mac":
        out = os.path.join(ROOT, "dist", "mac", "ChatFormatter-Mac.dotm")
    else:
        out = os.path.join(ROOT, "ChatFormatter.dotm")

    build(out)


if __name__ == "__main__":
    main()