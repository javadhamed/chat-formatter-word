"""Verify VBA linting, MS-OVBA compression, CFB packaging, and ChatFormatter.dotm."""

import base64
import hashlib
import os
import re
import stat
import struct
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "tools"))
sys.path.insert(0, os.path.join(ROOT, "tests"))

import build_dotm
import build_installers
import lint_vba
import vbabuild


def check_chunks(blob, label):
    assert blob and blob[0] == 0x01, f"{label}: missing 0x01 OVBA signature"
    pos = 1
    while pos < len(blob):
        header = struct.unpack_from("<H", blob, pos)[0]
        sig = (header >> 12) & 0x07
        size = (header & 0x0FFF) + 3
        assert sig == 0b011, f"{label}: bad chunk signature {sig:#x} at {pos}"
        assert pos + size <= len(blob), f"{label}: truncated chunk at {pos}"
        assert size <= 4096, f"{label}: chunk too big ({size} bytes)"
        pos += size
    assert pos == len(blob), f"{label}: trailing bytes after last chunk"


def check_cfb_directory(cfb_bytes):
    sec_sz = 1 << struct.unpack_from("<H", cfb_bytes, 30)[0]
    dir_start = struct.unpack_from("<I", cfb_bytes, 48)[0]
    num_fat = struct.unpack_from("<I", cfb_bytes, 44)[0]
    fat_secs = [s for s in struct.unpack_from("<109I", cfb_bytes, 76) if s < 0xFFFFFFFC][:num_fat]
    fat = []
    for s in fat_secs:
        off = (s + 1) * sec_sz
        fat.extend(struct.unpack("<" + "I" * (sec_sz // 4), cfb_bytes[off:off + sec_sz]))

    dir_buf = bytearray()
    s = dir_start
    while s < 0xFFFFFFFC:
        off = (s + 1) * sec_sz
        dir_buf.extend(cfb_bytes[off:off + sec_sz])
        s = fat[s]

    for i in range(0, len(dir_buf), 128):
        chunk = dir_buf[i:i + 128]
        kind = chunk[66]
        if kind == 1:  # storage (MS-CFB 2.6.3: start must be 0)
            start = struct.unpack_from("<I", chunk, 116)[0]
            assert start == 0, f"storage entry {i // 128} has start={start:#x}, expected 0"


def verify_dotm(path):
    with zipfile.ZipFile(path) as z:
        assert "word/vbaProject.bin" in z.namelist()
        assert "word/vbaData.xml" in z.namelist()
        vba_bin = z.read("word/vbaProject.bin")
        vba_data_xml = z.read("word/vbaData.xml")

    check_cfb_directory(vba_bin)
    streams = vbabuild.read_cfb(vba_bin)

    for req in ("PROJECT", "PROJECTwm", "VBA/dir", "VBA/_VBA_PROJECT",
                "VBA/ChatFormatter", "VBA/ThisDocument"):
        assert req in streams, f"missing stream {req}"

    assert streams["VBA/_VBA_PROJECT"] == b"\xcc\x61\xff\xff\x00\xff\xff", (
        f"unexpected _VBA_PROJECT header: {streams['VBA/_VBA_PROJECT']!r}"
    )

    check_chunks(streams["VBA/dir"], "VBA/dir")
    dir_plain = vbabuild.decompress(streams["VBA/dir"])
    offsets = vbabuild.parse_module_offsets(dir_plain)
    assert offsets == {"ThisDocument": 0, "ChatFormatter": 0}, (
        f"expected zeroed module offsets, got {offsets}"
    )

    for mod, fname in (("ChatFormatter", "ChatFormatter.bas"),
                       ("ThisDocument", "ThisDocument.cls")):
        check_chunks(streams["VBA/" + mod], "VBA/" + mod)
        got = vbabuild.decompress(streams["VBA/" + mod])
        want = build_dotm.read_source(fname)
        assert got == want, f"VBA/{mod} stream does not match src/{fname}"
        assert b"\n" not in got.replace(b"\r\n", b""), f"bare LF in VBA/{mod}"
        assert b"\r" not in got.replace(b"\r\n", b""), f"bare CR in VBA/{mod}"

    cf_text = vbabuild.decompress(streams["VBA/ChatFormatter"]).decode("latin-1")
    td_text = vbabuild.decompress(streams["VBA/ThisDocument"]).decode("latin-1")

    assert cf_text.startswith('Attribute VB_Name = "ChatFormatter"\r\n')
    assert cf_text.count("VB_Name") == 1
    assert not re.search(r"^ute\s+VB_Name", cf_text, re.M)
    assert 'CreateObject("WScript.Shell")' not in cf_text
    assert "HKCU\\Software\\ChatFormatter\\" not in cf_text

    assert td_text.startswith('Attribute VB_Name = "ThisDocument"\r\n')
    assert "VERSION 1.0 CLASS" not in td_text
    assert "Private WithEvents mApp As Word.Application" in td_text
    assert "mApp_WindowBeforeRightClick(ByVal Sel As Selection, Cancel As Boolean)" in td_text
    assert "Sub WindowBeforeRightClick(" not in td_text

    # Every macro listed in word/vbaData.xml must be a Public Sub in ChatFormatter.
    ns = {"wne": "http://schemas.microsoft.com/office/word/2006/wordml"}
    root = ET.fromstring(vba_data_xml)
    macros = [
        el.attrib["{http://schemas.microsoft.com/office/word/2006/wordml}name"].split(".")[-1]
        for el in root.findall(".//wne:mcd", ns)
    ]
    assert macros, "expected macro entries in word/vbaData.xml"
    for macro in macros:
        assert re.search(rf"^Public\s+Sub\s+{re.escape(macro)}\b", cf_text, re.M), (
            f"macro {macro} from vbaData.xml not found as Public Sub"
        )


def test_lint_catches_regressions():
    assert lint_vba.main([
        os.path.join(ROOT, "src", "ChatFormatter.bas"),
        os.path.join(ROOT, "src", "ThisDocument.cls"),
    ]) == 0

    with tempfile.NamedTemporaryFile("w", suffix=".bas", delete=False) as tf:
        tf.write(
            'Attribute VB_Name = "Broken"\n'
            'Option Explicit\n'
            'Public Sub Foo()\n'
            'End Sub\n'
            'ute VB_Name = "Broken"\n'
        )
        bad_path = tf.name
    try:
        mod = lint_vba.Module(bad_path)
        probs = mod.report()
        assert any("outside procedure" in p for p in probs), (
            f"expected outside-procedure lint failure, got {probs}"
        )
    finally:
        os.unlink(bad_path)


def test_rebuilt_dotm_matches_repo():
    dotm = os.path.join(ROOT, "ChatFormatter.dotm")
    verify_dotm(dotm)
    with tempfile.TemporaryDirectory() as tmp:
        rebuilt = os.path.join(tmp, "Rebuilt.dotm")
        build_dotm.build(rebuilt)
        verify_dotm(rebuilt)
        with zipfile.ZipFile(dotm) as z1, zipfile.ZipFile(rebuilt) as z2:
            assert z1.namelist() == z2.namelist()
            for name in z1.namelist():
                assert z1.read(name) == z2.read(name), f"mismatch in {name}"


def test_standalone_installers():
    dotm_path = os.path.join(ROOT, "ChatFormatter.dotm")
    with open(dotm_path, "rb") as fh:
        dotm_bytes = fh.read()
    expected_sha = hashlib.sha256(dotm_bytes).hexdigest()
    expected_size = len(dotm_bytes)

    win_path = os.path.join(ROOT, "ChatFormatter-Windows-Installer.bat")
    mac_path = os.path.join(ROOT, "ChatFormatter-Mac-Installer.command")
    zip_win_path = os.path.join(ROOT, "install.bat")
    zip_mac_path = os.path.join(ROOT, "packaging", "mac", "install.command")

    # Line ending and mode invariants
    for bat in (win_path, zip_win_path):
        with open(bat, "rb") as fh:
            raw = fh.read()
        assert b"\n" not in raw.replace(b"\r\n", b""), f"bare LF in {bat}"
        assert b"\r" not in raw.replace(b"\r\n", b""), f"bare CR in {bat}"

    for cmd in (mac_path, zip_mac_path):
        with open(cmd, "rb") as fh:
            raw = fh.read()
        assert b"\r" not in raw, f"CR found in {cmd}"
        assert os.stat(cmd).st_mode & stat.S_IXUSR, f"{cmd} missing +x mode"
        res = subprocess.run(["bash", "-n", cmd], capture_output=True, text=True)
        assert res.returncode == 0, f"bash -n failed for {cmd}: {res.stderr}"

    # Verify Windows standalone installer payload & PowerShell command integrity
    with open(win_path, "rb") as fh:
        win_raw = fh.read()
    assert win_raw == build_installers.render_windows_installer(dotm_bytes), (
        "ChatFormatter-Windows-Installer.bat is out of sync with ChatFormatter.dotm"
    )
    win_text = win_raw.decode("ascii")
    assert f'set "CF_SIZE={expected_size}"' in win_text
    assert f'set "CF_SHA256={expected_sha}"' in win_text
    ps_lines = [ln for ln in win_text.splitlines() if "powershell -NoProfile" in ln]
    assert len(ps_lines) == 1 and "-Command" in ps_lines[0] and "[regex]::Matches" in ps_lines[0]

    m_win = re.search(
        r"^-----BEGIN CHATFORMATTER DOTM-----\r\n(.*?)\r\n-----END CHATFORMATTER DOTM-----\r\n?$",
        win_text,
        re.M | re.S,
    )
    assert m_win, "missing payload markers in ChatFormatter-Windows-Installer.bat"
    assert base64.b64decode(m_win.group(1)) == dotm_bytes

    # Verify macOS standalone installer matches template and runs end-to-end
    with open(mac_path, "rb") as fh:
        mac_raw = fh.read()
    assert mac_raw == build_installers.render_mac_installer(dotm_bytes), (
        "ChatFormatter-Mac-Installer.command is out of sync with ChatFormatter.dotm"
    )

    with tempfile.TemporaryDirectory() as fake_home:
        old_dirs = [
            os.path.join(fake_home, "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"),
            os.path.join(fake_home, "Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word"),
            os.path.join(fake_home, "Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized"),
            os.path.join(fake_home, "Library/Application Support/Microsoft/Office/Word/STARTUP"),
            os.path.join(fake_home, "Documents/Microsoft User Data/Word Startup"),
        ]
        for idx, d in enumerate(old_dirs):
            os.makedirs(d, exist_ok=True)
            with open(os.path.join(d, f"ChatFormatter-old-{idx}.dotm"), "wb") as f:
                f.write(b"stale")
            with open(os.path.join(d, "ChatFormatter.dotm"), "wb") as f:
                f.write(b"old-version")
            with open(os.path.join(d, "~$ChatFormatter.dotm"), "wb") as f:
                f.write(b"lock")

        old_cfg = os.path.join(fake_home, ".chatformatter")
        os.makedirs(old_cfg, exist_ok=True)
        with open(os.path.join(old_cfg, "settings.ini"), "w") as f:
            f.write("AutoFormat = 0\nStaleKey = 999\n")

        env = os.environ.copy()
        env["HOME"] = fake_home
        env["CF_NO_PAUSE"] = "1"
        env["CF_SKIP_WORD_CHECK"] = "1"
        res = subprocess.run(["bash", mac_path], env=env, capture_output=True, text=True)
        assert res.returncode == 0, f"Mac installer failed: {res.stdout}\n{res.stderr}"

        # Every old/lock copy must be gone; only the fresh ChatFormatter.dotm exists in target
        target = os.path.join(old_dirs[0], "ChatFormatter.dotm")
        assert os.path.exists(target)
        with open(target, "rb") as f:
            assert f.read() == dotm_bytes
        for idx, d in enumerate(old_dirs):
            assert not os.path.exists(os.path.join(d, f"ChatFormatter-old-{idx}.dotm"))
            assert not os.path.exists(os.path.join(d, "~$ChatFormatter.dotm"))
            if idx > 0:
                assert not os.path.exists(os.path.join(d, "ChatFormatter.dotm"))

        with open(os.path.join(old_cfg, "settings.ini"), "r") as f:
            assert f.read() == "AutoFormat = 1\nTables = 1\n"


def test_versions_match():
    with open(os.path.join(ROOT, "src", "ChatFormatter.bas"), "r", encoding="utf-8") as fh:
        bas = fh.read()
    m = re.search(r'^Public Const CF_VERSION As String = "([^"]+)"', bas, re.M)
    assert m, "CF_VERSION not found in ChatFormatter.bas"
    ver = m.group(1)

    import package_mac
    import package_win
    assert package_win.VERSION == ver, f"package_win.VERSION={package_win.VERSION} != {ver}"
    assert package_mac.VERSION == ver, f"package_mac.VERSION={package_mac.VERSION} != {ver}"


def main():
    test_lint_catches_regressions()
    print("  PASS  lint_vba catches stray outside-procedure code and passes src/")
    test_rebuilt_dotm_matches_repo()
    print("  PASS  ChatFormatter.dotm CFB/MS-OVBA/sources verified and reproducible")
    test_standalone_installers()
    print("  PASS  Windows & macOS standalone installers verified and tested end-to-end")
    test_versions_match()
    print("  PASS  package versions match CF_VERSION")
    return 0


if __name__ == "__main__":
    sys.exit(main())
