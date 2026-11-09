"""Generate standalone self-contained installers for Windows and macOS.

Run:  python3 tools/build_installers.py

Reads ChatFormatter.dotm and emits two self-contained single-file installers
in the repository root:
  - ChatFormatter-Windows-Installer.bat
  - ChatFormatter-Mac-Installer.command

Each installer embeds ChatFormatter.dotm directly (verified by byte size and
SHA-256), automatically removes any previous installed version first, and
then installs the new version cleanly.
"""

import base64
import hashlib
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_DOTM = os.path.join(ROOT, "ChatFormatter.dotm")
WIN_OUT = os.path.join(ROOT, "ChatFormatter-Windows-Installer.bat")
MAC_OUT = os.path.join(ROOT, "ChatFormatter-Mac-Installer.command")

BEGIN_MARKER = "-----BEGIN CHATFORMATTER DOTM-----"
END_MARKER = "-----END CHATFORMATTER DOTM-----"


def _wrap_b64(raw_bytes, width=76):
    b64 = base64.b64encode(raw_bytes).decode("ascii")
    return [b64[i:i + width] for i in range(0, len(b64), width)]


def render_windows_installer(dotm_bytes):
    size = len(dotm_bytes)
    sha256 = hashlib.sha256(dotm_bytes).hexdigest()
    lines_b64 = "\r\n".join(_wrap_b64(dotm_bytes))

    script = f"""@echo off
setlocal EnableExtensions

REM ==============================================================
REM   Chat Formatter for Word - Standalone Windows Installer
REM   Self-contained: embeds ChatFormatter.dotm, removes any old
REM   installation first, then installs the new version.
REM ==============================================================

title Chat Formatter - Install / Upgrade

set "TEMPLATE_NAME=ChatFormatter.dotm"
set "CF_SELF=%~f0"
set "CF_SIZE={size}"
set "CF_SHA256={sha256}"
set "CF_TMP_DOTM=%TEMP%\\ChatFormatter_new_%RANDOM%_%RANDOM%.dotm"
set "CF_B64_TXT=%TEMP%\\ChatFormatter_b64_%RANDOM%_%RANDOM%.txt"
set "CF_EXIT=0"

echo.
echo ==============================================================
echo   Chat Formatter for Word - Standalone Installer (Windows)
echo ==============================================================
echo.
echo   User: %USERNAME%

REM --- 1. Extract & verify embedded ChatFormatter.dotm ------------
echo [1/5] Preparing new ChatFormatter package...
call :ExtractPayload
if errorlevel 1 (
    echo [ERROR] Failed to extract or verify the embedded template.
    echo         Make sure you downloaded the real .bat file ^(not an HTML page^).
    set "CF_EXIT=1"
    goto :Cleanup
)
echo       OK ^(verified %CF_SIZE% bytes^)

REM --- 2. Ensure Word is closed before touching STARTUP -----------
echo [2/5] Checking if Microsoft Word is running...
call :EnsureWordClosed
if errorlevel 1 (
    set "CF_EXIT=1"
    goto :Cleanup
)

REM --- 3. Remove previous version completely ----------------------
echo [3/5] Removing any previous version...
set "DEFAULT_STARTUP=%APPDATA%\\Microsoft\\Word\\STARTUP"
set "TARGET_STARTUP=%DEFAULT_STARTUP%"
set "ACTIVE_VER="
set "REG_STARTUP="

call :CleanDir "%DEFAULT_STARTUP%"
for %%V in (16.0 15.0 14.0) do (
    reg query "HKCU\\Software\\Microsoft\\Office\\%%V\\Word" >nul 2>&1
    if not errorlevel 1 (
        call :CheckVersionStartup "%%V"
    )
)

set "CFHOME=%USERPROFILE%"
if not defined CFHOME set "CFHOME=%HOMEDRIVE%%HOMEPATH%"

call :CleanDir "%LOCALAPPDATA%\\Microsoft\\Word\\STARTUP"
call :CleanDir "%LOCALAPPDATA%\\Microsoft\\Office\\Word\\STARTUP"
call :CleanDir "%CFHOME%\\Documents\\Custom Office Templates"

REM Remove legacy v1.0.0 registry key and stale trusted locations
reg delete "HKCU\\Software\\ChatFormatter" /f >nul 2>&1
for %%V in (14.0 15.0 16.0) do (
    reg delete "HKCU\\Software\\Microsoft\\Office\\%%V\\Word\\Security\\Trusted Locations\\LocationChatFormatter" /f >nul 2>&1
)

if exist "%CFHOME%\\.chatformatter" (
    rmdir /S /Q "%CFHOME%\\.chatformatter" >nul 2>&1
)
echo       Previous files and settings cleaned.

REM --- 4. Install new version -------------------------------------
echo [4/5] Installing new version to Word STARTUP...
echo       Target folder: %TARGET_STARTUP%
if not exist "%TARGET_STARTUP%" mkdir "%TARGET_STARTUP%" 2>nul
if not exist "%TARGET_STARTUP%" (
    echo       [WARN] Could not access %TARGET_STARTUP%, falling back to default STARTUP.
    set "TARGET_STARTUP=%DEFAULT_STARTUP%"
    if not exist "%TARGET_STARTUP%" mkdir "%TARGET_STARTUP%" 2>nul
)
if not exist "%TARGET_STARTUP%" (
    echo [ERROR] Could not create STARTUP folder: %TARGET_STARTUP%
    set "CF_EXIT=1"
    goto :Cleanup
)
copy /Y "%CF_TMP_DOTM%" "%TARGET_STARTUP%\\%TEMPLATE_NAME%" >nul
if errorlevel 1 (
    echo [ERROR] Copy failed. Close all Word windows and run this again.
    set "CF_EXIT=1"
    goto :Cleanup
)
set "CF_INST_SIZE="
for %%F in ("%TARGET_STARTUP%\\%TEMPLATE_NAME%") do set "CF_INST_SIZE=%%~zF"
if not "%CF_INST_SIZE%"=="%CF_SIZE%" (
    echo [ERROR] Installed file verification failed at:
    echo         %TARGET_STARTUP%\\%TEMPLATE_NAME%
    set "CF_EXIT=1"
    goto :Cleanup
)
echo       OK -^> %TARGET_STARTUP%\\%TEMPLATE_NAME% ^(%CF_INST_SIZE% bytes^)

REM --- 5. Register trusted location and write fresh settings ------
echo [5/5] Registering trusted location and default settings...
set "REG_OK=0"
for %%V in (14.0 15.0 16.0) do (
    reg query "HKCU\\Software\\Microsoft\\Office\\%%V\\Word" >nul 2>&1
    if not errorlevel 1 (
        call :RegisterTrust "%%V" "%TARGET_STARTUP%"
    )
)
if "%REG_OK%"=="0" call :RegisterTrust "16.0" "%TARGET_STARTUP%"
if "%REG_OK%"=="1" (
    echo       Trusted location registered.
) else (
    echo       [WARN] Could not register Trusted Location automatically.
    echo        In Word: File ^> Options ^> Trust Center ^> Trust Center Settings
    echo        ^> Trusted Locations ^> Add "%TARGET_STARTUP%"
)

if not exist "%CFHOME%\\.chatformatter" mkdir "%CFHOME%\\.chatformatter" 2>nul
(
    echo AutoFormat = 1
    echo Tables = 1
) > "%CFHOME%\\.chatformatter\\settings.ini"
if exist "%CFHOME%\\.chatformatter\\settings.ini" (
    echo       Created %CFHOME%\\.chatformatter\\settings.ini
)

echo.
echo ==============================================================
echo   Installation complete!
echo ==============================================================
echo.
echo   Installed: %TARGET_STARTUP%\\%TEMPLATE_NAME%
echo.
echo   How to use:
echo     1. Open Word
echo     2. Copy text from an AI chat ^(ChatGPT, DeepSeek, Gemini, ...^)
echo     3. Paste into Word, select it, and RIGHT-CLICK
echo.
echo   Manual macros ^(Alt+F8^): FormatChatText, ToggleAutoFormat, ShowSettings
echo.
goto :Cleanup

:ExtractPayload
if exist "%CF_TMP_DOTM%" del /F /Q /A "%CF_TMP_DOTM%" >nul 2>&1
where powershell >nul 2>&1
if not errorlevel 1 (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$t=[IO.File]::ReadAllText($env:CF_SELF);$b1='-----BEGIN '+'CHATFORMATTER DOTM-----';$e1='-----END '+'CHATFORMATTER DOTM-----';$pat='(?ms)^'+[regex]::Escape($b1)+'\\r?\\n(.*?)\\r?\\n'+[regex]::Escape($e1)+'(?:\\r?\\n|$)';$ms=[regex]::Matches($t,$pat);if($ms.Count -eq 0){{exit 2}};$b=[Convert]::FromBase64String(($ms[$ms.Count-1].Groups[1].Value -replace '[^A-Za-z0-9+/=]',''));if($b.Length -ne [int]$env:CF_SIZE){{exit 3}};$h=[BitConverter]::ToString([Security.Cryptography.SHA256]::Create().ComputeHash($b)).Replace('-','').ToLower();if($h -ne $env:CF_SHA256){{exit 4}};[IO.File]::WriteAllBytes($env:CF_TMP_DOTM,$b)" >nul 2>&1
    if not errorlevel 1 (
        if exist "%CF_TMP_DOTM%" exit /b 0
    )
    echo       PowerShell decode skipped; trying certutil fallback...
)

REM Fallback when PowerShell is unavailable or restricted
> "%CF_B64_TXT%" echo -----BEGIN CERTIFICATE-----
findstr /R "^[A-Za-z0-9+/][A-Za-z0-9+/=]*$" "%CF_SELF%" >> "%CF_B64_TXT%"
>> "%CF_B64_TXT%" echo -----END CERTIFICATE-----

certutil -f -decode "%CF_B64_TXT%" "%CF_TMP_DOTM%" >nul 2>&1
if errorlevel 1 exit /b 1
set "CF_GOT_SIZE="
for %%F in ("%CF_TMP_DOTM%") do set "CF_GOT_SIZE=%%~zF"
if not "%CF_GOT_SIZE%"=="%CF_SIZE%" exit /b 1

set "CF_HASH_LINE="
for /f "skip=1 delims=" %%H in ('certutil -hashfile "%CF_TMP_DOTM%" SHA256 2^>nul') do (
    if not defined CF_HASH_LINE set "CF_HASH_LINE=%%H"
)
if not defined CF_HASH_LINE exit /b 0
set "CF_HASH_LINE=%CF_HASH_LINE: =%"
echo "%CF_HASH_LINE%" | findstr /R "^\"[0-9A-Fa-f][0-9A-Fa-f]*\"$" >nul || exit /b 0
if /I not "%CF_HASH_LINE%"=="%CF_SHA256%" exit /b 1
exit /b 0

:EnsureWordClosed
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if errorlevel 1 (
    echo       Word is not running.
    exit /b 0
)
echo       Word is running. Requesting graceful close...
taskkill /IM WINWORD.EXE >nul 2>&1
for /L %%I in (1,1,10) do (
    ping -n 2 127.0.0.1 >nul
    tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
    if errorlevel 1 (
        echo       Word closed.
        exit /b 0
    )
)
echo       Word is still open ^(you may have an unsaved document^).
:WordPrompt
set "CF_ANS="
set /p "CF_ANS=      Save/close Word then press Enter to retry, [F] to force-close, or [Q] to quit: "
if /I "%CF_ANS%"=="Q" exit /b 1
if /I "%CF_ANS%"=="F" (
    taskkill /F /IM WINWORD.EXE >nul 2>&1
    ping -n 2 127.0.0.1 >nul
)
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 goto :WordPrompt
echo       Word closed.
exit /b 0

:CheckVersionStartup
set "CF_VER=%~1"
set "CF_IS_ACTIVE="
if not defined ACTIVE_VER (
    set "ACTIVE_VER=%CF_VER%"
    set "CF_IS_ACTIVE=1"
)
set "CF_CAND="
for /f "tokens=2,*" %%A in ('reg query "HKCU\\Software\\Microsoft\\Office\\%CF_VER%\\Word\\Options" /v STARTUP-PATH 2^>nul ^| find /I "STARTUP-PATH"') do (
    call set "CF_CAND=%%B"
)
if not defined CF_CAND exit /b 0
set "CF_CAND=%CF_CAND:"=%"
if "%CF_CAND%"=="" exit /b 0
if "%CF_CAND:~-1%"=="\\" set "CF_CAND=%CF_CAND:~0,-1%"
if "%CF_CAND%"=="" exit /b 0
call :CleanDir "%CF_CAND%"
if not defined CF_IS_ACTIVE exit /b 0
if defined REG_STARTUP exit /b 0
if "%CF_CAND:~1,1%"==":" (
    if not exist "%CF_CAND:~0,3%" exit /b 0
)
set "REG_STARTUP=%CF_CAND%"
set "TARGET_STARTUP=%CF_CAND%"
exit /b 0

:CleanDir
set "CF_DIR=%~1"
if "%CF_DIR%"=="" exit /b 0
if exist "%CF_DIR%\\ChatFormatter*.dotm" del /F /Q /A "%CF_DIR%\\ChatFormatter*.dotm" >nul 2>&1
if exist "%CF_DIR%\\~$ChatFormatter*.dotm" del /F /Q /A "%CF_DIR%\\~$ChatFormatter*.dotm" >nul 2>&1
exit /b 0

:RegisterTrust
set "SECKEY=HKCU\\Software\\Microsoft\\Office\\%~1\\Word\\Security\\Trusted Locations"
reg query "%SECKEY%" >nul 2>&1
if errorlevel 1 reg add "%SECKEY%" /f >nul 2>&1
set "CF_TPATH=%~2"
if "%CF_TPATH:~0,2%"=="\\\\" (
    reg add "%SECKEY%" /v AllowNetworkLocations /t REG_DWORD /d 1 /f >nul 2>&1
)
reg add "%SECKEY%\\LocationChatFormatter" /v Path /t REG_EXPAND_SZ /d "%~2" /f >nul 2>&1
reg add "%SECKEY%\\LocationChatFormatter" /v Description /t REG_SZ /d "Chat Formatter for Word" /f >nul 2>&1
reg add "%SECKEY%\\LocationChatFormatter" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul 2>&1
if not errorlevel 1 set "REG_OK=1"
exit /b 0

:Cleanup
if exist "%CF_TMP_DOTM%" del /F /Q /A "%CF_TMP_DOTM%" >nul 2>&1
if exist "%CF_B64_TXT%" del /F /Q /A "%CF_B64_TXT%" >nul 2>&1
if not defined CF_NO_PAUSE pause
endlocal & exit /b %CF_EXIT%

{BEGIN_MARKER}
{lines_b64}
{END_MARKER}
"""
    # Normalize all line endings strictly to CRLF for cmd.exe.
    script = script.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")
    return script.encode("ascii")


def render_mac_installer(dotm_bytes):
    size = len(dotm_bytes)
    sha256 = hashlib.sha256(dotm_bytes).hexdigest()
    lines_b64 = "\n".join(_wrap_b64(dotm_bytes))

    script = f"""#!/bin/bash
# ==============================================================
#   Chat Formatter for Word - Standalone macOS Installer
#   Self-contained: embeds ChatFormatter.dotm, removes any old
#   installation first, then installs the new version.
# ==============================================================
set -uo pipefail

TEMPLATE="ChatFormatter.dotm"
EXPECTED_SIZE={size}
EXPECTED_SHA256="{sha256}"

STARTUP="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
LEGACY_STARTUPS=(
    "$STARTUP"
    "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word"
    "$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Templates.localized"
    "$HOME/Library/Application Support/Microsoft/Office/Word/STARTUP"
    "$HOME/Documents/Microsoft User Data/Word Startup"
)
SETTINGS_DIR="$HOME/.chatformatter"
SETTINGS="$SETTINGS_DIR/settings.ini"

TMP_B64=""
TMP_DOTM=""

cleanup() {{
    [ -n "$TMP_B64" ] && rm -f "$TMP_B64"
    [ -n "$TMP_DOTM" ] && rm -f "$TMP_DOTM"
}}
trap cleanup EXIT

say() {{ printf '%s\\n' "$*"; }}
pause_exit() {{
    local code="$1"
    if [ "${{CF_NO_PAUSE:-0}}" != "1" ]; then
        printf '\\nPress any key to close...'
        read -rsn1 || true
        printf '\\n'
    fi
    exit "$code"
}}
die() {{
    printf '[ERROR] %s\\n' "$*" >&2
    pause_exit 1
}}

calc_sha256() {{
    local file="$1"
    if command -v shasum >/dev/null 2>&1; then
        shasum -a 256 "$file" | awk '{{print $1}}'
    elif command -v sha256sum >/dev/null 2>&1; then
        sha256sum "$file" | awk '{{print $1}}'
    elif command -v openssl >/dev/null 2>&1; then
        openssl dgst -sha256 "$file" | awk '{{print $NF}}'
    else
        return 1
    fi
}}

decode_b64() {{
    local src="$1"
    local dst="$2"
    if base64 -D < "$src" > "$dst" 2>/dev/null; then
        return 0
    fi
    if base64 -d < "$src" > "$dst" 2>/dev/null; then
        return 0
    fi
    if command -v openssl >/dev/null 2>&1 && openssl base64 -d -in "$src" -out "$dst" 2>/dev/null; then
        return 0
    fi
    return 1
}}

extract_payload() {{
    local base_tmp="${{TMPDIR:-/tmp}}"
    base_tmp="${{base_tmp%/}}"
    TMP_B64="$(mktemp "$base_tmp/cf_b64.XXXXXX")" || return 1
    TMP_DOTM="$(mktemp "$base_tmp/cf_dotm.XXXXXX")" || return 1

    awk '
        BEGIN {{ b = "-----BEGIN " "CHATFORMATTER DOTM-----"; e = "-----END " "CHATFORMATTER DOTM-----" }}
        $0 == b {{ flag=1; next }}
        $0 == e {{ flag=0; exit }}
        flag {{ print }}
    ' "$0" > "$TMP_B64"

    [ -s "$TMP_B64" ] || return 1
    decode_b64 "$TMP_B64" "$TMP_DOTM" || return 1

    local got_size got_sha
    got_size="$(wc -c < "$TMP_DOTM" | tr -d ' ')"
    [ "$got_size" = "$EXPECTED_SIZE" ] || return 1
    got_sha="$(calc_sha256 "$TMP_DOTM")" || return 1
    [ "$got_sha" = "$EXPECTED_SHA256" ] || return 1
    return 0
}}

say "=============================================================="
say "  Chat Formatter for Word - Standalone Installer (macOS)"
say "=============================================================="
say ""

# --- 1. Extract and verify embedded package --------------------
say "[1/5] Preparing new ChatFormatter package..."
extract_payload || die "Could not extract or verify the embedded template."
say "      OK (verified $EXPECTED_SIZE bytes)"

# --- 2. Ensure Word is closed ----------------------------------
if [ "${{CF_SKIP_WORD_CHECK:-0}}" != "1" ] && command -v pgrep >/dev/null 2>&1 && pgrep -x "Microsoft Word" >/dev/null 2>&1; then
    say "[2/5] Microsoft Word is running. Requesting quit..."
    if command -v osascript >/dev/null 2>&1; then
        osascript -e 'tell application "Microsoft Word" to quit' >/dev/null 2>&1 || true
    fi
    for _ in $(seq 1 30); do
        pgrep -x "Microsoft Word" >/dev/null 2>&1 || break
        sleep 0.5
    done
    while pgrep -x "Microsoft Word" >/dev/null 2>&1; do
        say "      Word is still open (you may have an unsaved document)."
        printf '      Save and quit Word, then press Enter to continue (or q to cancel): '
        read -r ans || die "Cancelled."
        if [ "$ans" = "q" ] || [ "$ans" = "Q" ]; then
            die "Cancelled by user."
        fi
    done
    say "      Word closed."
else
    say "[2/5] Word is not running."
fi

# --- 3. Remove previous version completely ---------------------
say "[3/5] Removing any previous version..."
for dir in "${{LEGACY_STARTUPS[@]}}"; do
    if [ -d "$dir" ]; then
        rm -f "$dir"/ChatFormatter*.dotm "$dir"/~\\$ChatFormatter*.dotm 2>/dev/null || true
    fi
done
if [ -d "$SETTINGS_DIR" ]; then
    rm -rf "$SETTINGS_DIR" || die "Could not remove old settings at $SETTINGS_DIR"
fi
say "      Previous files and settings cleaned."

# --- 4. Install new version ------------------------------------
say "[4/5] Installing new version into Word STARTUP folder..."
mkdir -p "$STARTUP" || die "Could not create $STARTUP"
cp "$TMP_DOTM" "$STARTUP/$TEMPLATE" || die "Copy failed. Check folder permissions for $STARTUP."
chmod 644 "$STARTUP/$TEMPLATE" 2>/dev/null || true
if command -v xattr >/dev/null 2>&1; then
    xattr -d com.apple.quarantine "$STARTUP/$TEMPLATE" 2>/dev/null || true
fi

installed_sha="$(calc_sha256 "$STARTUP/$TEMPLATE" || true)"
[ "$installed_sha" = "$EXPECTED_SHA256" ] || die "Installed file checksum verification failed."
say "      OK -> $STARTUP/$TEMPLATE"

# --- 5. Write fresh default settings ---------------------------
say "[5/5] Writing default settings..."
mkdir -p "$SETTINGS_DIR" || die "Could not create $SETTINGS_DIR"
printf 'AutoFormat = 1\\nTables = 1\\n' > "$SETTINGS" || die "Could not write $SETTINGS"
say "      Created $SETTINGS"

say ""
say "=============================================================="
say "  Installation complete!"
say "=============================================================="
say ""
say "  Installed: $STARTUP/$TEMPLATE"
say ""
say "  First launch note:"
say "    When Word opens, choose Yes if asked to enable macros."
say "    (Word > Settings > Security > Trusted Locations if needed.)"
say ""
say "  How to use:"
say "    1. Copy an answer from ChatGPT, DeepSeek, Gemini, ..."
say "    2. Paste it into Word"
say "    3. Select the pasted text and RIGHT-CLICK"
say ""
say "  Manual macros (Tools > Macro): FormatChatText, ToggleAutoFormat, ShowSettings"

pause_exit 0

{BEGIN_MARKER}
{lines_b64}
{END_MARKER}
"""
    script = script.replace("\r\n", "\n").replace("\r", "\n")
    return script.encode("ascii")


def build_all(dotm_path=DEFAULT_DOTM, win_out=WIN_OUT, mac_out=MAC_OUT):
    if not os.path.exists(dotm_path):
        sys.exit("Missing template: %s" % dotm_path)
    with open(dotm_path, "rb") as fh:
        dotm_bytes = fh.read()

    win_bytes = render_windows_installer(dotm_bytes)
    with open(win_out, "wb") as fh:
        fh.write(win_bytes)
    print("Wrote %s (%d bytes)" % (win_out, len(win_bytes)))

    mac_bytes = render_mac_installer(dotm_bytes)
    with open(mac_out, "wb") as fh:
        fh.write(mac_bytes)
    os.chmod(mac_out, 0o755)
    print("Wrote %s (%d bytes, mode 0755)" % (mac_out, len(mac_bytes)))


def main():
    dotm = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_DOTM
    build_all(dotm_path=dotm)
    return 0


if __name__ == "__main__":
    sys.exit(main())
