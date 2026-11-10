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
        flag = (header >> 15) & 1
        size = (header & 0x0FFF) + 3
        assert sig == 0b011, f"{label}: bad chunk signature {sig:#x} at {pos}"
        assert pos + size <= len(blob), f"{label}: truncated chunk at {pos}"
        assert size <= 4098 if flag == 0 else size <= 4096, (
            f"{label}: chunk too big ({size} bytes, flag={flag})"
        )
        plain = vbabuild.decompress(b"\x01" + blob[pos:pos + size])
        is_last = (pos + size == len(blob))
        if flag == 0:
            assert size == 4098 and len(plain) == 4096, (
                f"{label}: raw chunk must be 4096 bytes, got {len(plain)}"
            )
        elif not is_last:
            assert len(plain) == 4096, (
                f"{label}: non-final compressed chunk decompressed to {len(plain)} != 4096 (MS-OVBA 2.4.1.1.4)"
            )
        else:
            assert len(plain) <= 4096, f"{label}: final chunk decompressed to {len(plain)} > 4096"
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

    entries = []
    for i in range(0, len(dir_buf), 128):
        chunk = dir_buf[i:i + 128]
        kind = chunk[66]
        if kind == 0:
            entries.append(None)
            continue
        nlen = struct.unpack_from("<H", chunk, 64)[0]
        name = chunk[:max(0, nlen - 2)].decode("utf-16-le", "replace")
        color = chunk[67]
        left, right, child = struct.unpack_from("<III", chunk, 68)
        start = struct.unpack_from("<I", chunk, 116)[0]
        if kind == 1:  # storage (MS-CFB 2.6.3: start must be 0)
            assert start == 0, f"storage entry {i // 128} has start={start:#x}, expected 0"
        entries.append((name, kind, color, left, right, child))

    def validate_rbtree(root_idx):
        if root_idx == vbabuild.NOSTREAM:
            return
        assert entries[root_idx][2] == 1, f"RB root {entries[root_idx][0]} must be black"

        def dfs(idx, parent_red):
            if idx == vbabuild.NOSTREAM:
                return 1, []
            name, kind, color, left, right, child = entries[idx]
            assert color in (0, 1), f"invalid color {color} for {name}"
            if parent_red:
                assert color == 1, f"adjacent red nodes at {name}"
            bh_l, keys_l = dfs(left, color == 0)
            bh_r, keys_r = dfs(right, color == 0)
            assert bh_l == bh_r, f"unequal black-height at {name}: {bh_l} != {bh_r}"
            key = (len(name), name.upper())
            if keys_l:
                assert keys_l[-1] < key, f"BST violation left of {name}"
            if keys_r:
                assert key < keys_r[0], f"BST violation right of {name}"
            if kind == 1:
                validate_rbtree(child)
            return bh_l + (1 if color == 1 else 0), keys_l + [key] + keys_r

        dfs(root_idx, False)

    for ent in entries:
        if ent and ent[1] == 5:
            validate_rbtree(ent[5])
            break


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
    assert "IsSetextContinuation(" in cf_text
    for removed_symbol in (
        "CF_RightClickHandler",
        "LooksLikeAIChat",
        "ToggleAutoFormat",
        "EnsureAppHook",
        "HasDashRunHint",
        "AutoFormat",
    ):
        assert removed_symbol not in cf_text, f"obsolete {removed_symbol} still in VBA/ChatFormatter"
    for required_symbol in (
        "BK_PIPETABLE",
        "ApplyPipeTable",
        "SplitSequentialNumbered",
        "NormalizeDashes",
        "CollapseDashesInRange",
        "DigitValue",
        "AppendParaSegments",
        "Font.BoldBi = True",
        "wdTableDirectionRtl",
    ):
        assert required_symbol in cf_text, f"missing {required_symbol} in VBA/ChatFormatter"
    m_scan = re.search(r"mBusy\s*=\s*True.*?nBlocks\s*=\s*ScanBlocks\(", cf_text, re.S)
    assert m_scan, "mBusy must be set before ScanBlocks so DoEvents cannot re-enter"

    assert td_text.startswith('Attribute VB_Name = "ThisDocument"\r\n')
    assert "VERSION 1.0 CLASS" not in td_text
    for removed_td in ("WithEvents", "WindowBeforeRightClick", "HookApplication", "CF_RightClickHandler"):
        assert removed_td not in td_text, f"obsolete {removed_td} still in VBA/ThisDocument"
    assert "Private Sub Document_Open()" in td_text
    assert "ChatFormatter.AutoExec" in td_text

    # Every macro listed in word/vbaData.xml must be a Public Sub in ChatFormatter.
    ns = {"wne": "http://schemas.microsoft.com/office/word/2006/wordml"}
    root = ET.fromstring(vba_data_xml)
    macros = [
        el.attrib["{http://schemas.microsoft.com/office/word/2006/wordml}name"].split(".")[-1]
        for el in root.findall(".//wne:mcd", ns)
    ]
    assert macros == [
        "FormatChatText",
        "FormatSelection",
        "ShowSettings",
        "AutoExec",
        "AutoOpen",
        "AutoSetup",
        "RemoveToolbar",
    ], f"unexpected vbaData.xml macros: {macros}"
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

    with tempfile.NamedTemporaryFile("w", suffix=".bas", delete=False) as tf:
        tf.write(
            'Attribute VB_Name = "ShortCircuit"\n'
            'Option Explicit\n'
            'Public Sub Foo()\n'
            '    Dim m As Long\n'
            '    Dim outKind() As Long\n'
            '    ReDim outKind(1 To 2)\n'
            '    If m > 0 And outKind(m) = 1 Then\n'
            '    End If\n'
            'End Sub\n'
        )
        sc_path = tf.name
    try:
        mod = lint_vba.Module(sc_path)
        probs = mod.report()
        assert any("short-circuit" in p for p in probs), (
            f"expected short-circuit lint failure, got {probs}"
        )
    finally:
        os.unlink(sc_path)


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

    # Verify findstr certutil fallback extracts ONLY the base64 payload lines
    before_payload = win_text.split("-----BEGIN CHATFORMATTER DOTM-----")[0]
    for idx, line in enumerate(before_payload.splitlines(), 1):
        assert not re.fullmatch(r"[A-Za-z0-9+/][A-Za-z0-9+/=]*", line), (
            f"script line {idx} ({line!r}) would be matched by certutil findstr fallback"
        )
    b64_lines = [
        ln for ln in win_text.splitlines()
        if re.fullmatch(r"[A-Za-z0-9+/][A-Za-z0-9+/=]*", ln)
    ]
    assert base64.b64decode("".join(b64_lines)) == dotm_bytes

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
            assert f.read() == "Tables = 1\n"

    # Docs must point both English & Persian users to Add-ins tab (never right-click)
    with open(os.path.join(ROOT, "README.md"), "r", encoding="utf-8") as f:
        readme = f.read()
    assert "Add-ins" in readme and "تب افزونه‌ها" in readme
    assert "ToggleAutoFormat" not in readme
    assert "WindowBeforeRightClick" not in readme

    with open(os.path.join(ROOT, "packaging", "mac", "README-mac.md"), "r", encoding="utf-8") as f:
        readme_mac = f.read()
    assert "Add-ins" in readme_mac and "تب افزونه‌ها" in readme_mac
    assert "ToggleAutoFormat" not in readme_mac


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


def test_ovba_compression_roundtrip():
    samples = [
        b"",
        b"x",
        b"A" * 5000,
        (b"Sub Foo()\r\n    MsgBox \"Hi\"\r\nEnd Sub\r\n") * 200,
        os.urandom(4096),  # exercises raw 4096-byte fallback chunk
    ]
    for idx, sample in enumerate(samples):
        comp = vbabuild.compress(sample)
        check_chunks(comp, f"sample[{idx}]")
        assert vbabuild.decompress(comp) == sample


def main():
    test_ovba_compression_roundtrip()
    print("  PASS  MS-OVBA 4096-byte LZ77 + raw fallback round-trips verified")
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
