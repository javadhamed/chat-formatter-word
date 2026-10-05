"""Assemble the downloadable macOS package.

Run:  python3 tools/package_mac.py

Produces dist/mac/ (the folder you can inspect) and dist/ChatFormatter-Mac-<v>.zip
(the file you hand to somebody).  The ZIP holds a single ChatFormatter-Mac/
folder so unzipping never scatters files across the Desktop.
"""

import os
import shutil
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VERSION = "1.0.0"

DIST = os.path.join(ROOT, "dist")
STAGE = os.path.join(DIST, "mac")
FOLDER = "ChatFormatter-Mac"
SAMPLE = os.path.expanduser("~/Downloads/لیست کامل و به.docx")

TEMPLATE = "ChatFormatter.dotm"


def build_template():
    sys.path.insert(0, os.path.join(ROOT, "tools"))
    import build_dotm

    out = os.path.join(STAGE, TEMPLATE)
    build_dotm.build(out)
    return out


def copy_if(path, dest, required=True):
    if not os.path.exists(path):
        if required:
            raise SystemExit("missing: " + path)
        return False
    shutil.copyfile(path, dest)
    return True


def main():
    if os.path.isdir(STAGE):
        shutil.rmtree(STAGE)
    os.makedirs(STAGE)

    print("building template...")
    template = build_template()

    for name in ("install.command", "uninstall.command", "README-mac.md"):
        copy_if(os.path.join(ROOT, "packaging", "mac", name), os.path.join(STAGE, name))
    for name in ("install.command", "uninstall.command"):
        os.chmod(os.path.join(STAGE, name), 0o755)
    copy_if(os.path.join(ROOT, "README.md"), os.path.join(STAGE, "README.md"))
    copy_if(os.path.join(ROOT, "LICENSE"), os.path.join(STAGE, "LICENSE"))
    have_sample = copy_if(SAMPLE, os.path.join(STAGE, "sample.docx"), required=False)
    if not have_sample:
        print("note: sample.docx not found, packaging without it")

    zip_path = os.path.join(DIST, "ChatFormatter-Mac-%s.zip" % VERSION)
    if os.path.exists(zip_path):
        os.remove(zip_path)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in sorted(os.listdir(STAGE)):
            z.write(os.path.join(STAGE, name), os.path.join(FOLDER, name))

    # The template is platform neutral, so keep the historical top-level copy
    # in step with what ships.
    shutil.copyfile(template, os.path.join(ROOT, TEMPLATE))

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