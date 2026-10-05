"""Assemble the downloadable Windows package.

Run:  python3 tools/package_win.py

The template itself is platform neutral - the same .dotm is shipped to both
platforms - so the build goes through the Python writer and needs neither
Word nor PowerShell.  Only the two .bat helpers are Windows specific.

Produces dist/win/ (the folder you can inspect) and
dist/ChatFormatter-Win-<v>.zip (the file you hand to somebody).
"""

import os
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VERSION = "1.0.0"

DIST = os.path.join(ROOT, "dist")
STAGE = os.path.join(DIST, "win")
FOLDER = "ChatFormatter-Win"

TEMPLATE = "ChatFormatter.dotm"
STAGED = ("install.bat", "uninstall.bat", "README.md", "LICENSE")


def build_template():
    sys.path.insert(0, os.path.join(ROOT, "tools"))
    import build_dotm

    out = os.path.join(STAGE, TEMPLATE)
    build_dotm.build(out)
    return out


def main():
    if os.path.isdir(STAGE):
        shutil.rmtree(STAGE)
    os.makedirs(STAGE)

    print("building template...")
    template = build_template()

    for name in STAGED:
        src = os.path.join(ROOT, name)
        if not os.path.exists(src):
            raise SystemExit("missing: " + src)
        shutil.copyfile(src, os.path.join(STAGE, name))

    # Keep the historical top-level copy in step with what ships.
    shutil.copyfile(template, os.path.join(ROOT, TEMPLATE))

    zip_path = os.path.join(DIST, "ChatFormatter-Win-%s.zip" % VERSION)
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(os.listdir(STAGE)):
            z.write(os.path.join(STAGE, name), os.path.join(FOLDER, name))

    print()
    print("package ready")
    print("  template  %8d bytes  %s" % (os.path.getsize(template), template))
    print("  archive   %8d bytes  %s" % (os.path.getsize(zip_path), zip_path))
    print()
    for name in sorted(os.listdir(STAGE)):
        print("  %-20s %8d" % (name, os.path.getsize(os.path.join(STAGE, name))))
    return 0


if __name__ == "__main__":
    sys.exit(main())