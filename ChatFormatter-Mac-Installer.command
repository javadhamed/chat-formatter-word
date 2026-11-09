#!/bin/bash
# ==============================================================
#   Chat Formatter for Word - Standalone macOS Installer
#   Self-contained: embeds ChatFormatter.dotm, removes any old
#   installation first, then installs the new version.
# ==============================================================
set -uo pipefail

TEMPLATE="ChatFormatter.dotm"
EXPECTED_SIZE=31047
EXPECTED_SHA256="9b511a2e8a30e12a01efaee29677359ef69322a3d5aebc560df4072ab159e306"

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

cleanup() {
    [ -n "$TMP_B64" ] && rm -f "$TMP_B64"
    [ -n "$TMP_DOTM" ] && rm -f "$TMP_DOTM"
}
trap cleanup EXIT

say() { printf '%s\n' "$*"; }
pause_exit() {
    local code="$1"
    if [ "${CF_NO_PAUSE:-0}" != "1" ]; then
        printf '\nPress any key to close...'
        read -rsn1 || true
        printf '\n'
    fi
    exit "$code"
}
die() {
    printf '[ERROR] %s\n' "$*" >&2
    pause_exit 1
}

calc_sha256() {
    local file="$1"
    if command -v shasum >/dev/null 2>&1; then
        shasum -a 256 "$file" | awk '{print $1}'
    elif command -v sha256sum >/dev/null 2>&1; then
        sha256sum "$file" | awk '{print $1}'
    elif command -v openssl >/dev/null 2>&1; then
        openssl dgst -sha256 "$file" | awk '{print $NF}'
    else
        return 1
    fi
}

decode_b64() {
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
}

extract_payload() {
    local base_tmp="${TMPDIR:-/tmp}"
    base_tmp="${base_tmp%/}"
    TMP_B64="$(mktemp "$base_tmp/cf_b64.XXXXXX")" || return 1
    TMP_DOTM="$(mktemp "$base_tmp/cf_dotm.XXXXXX")" || return 1

    awk '
        BEGIN { b = "-----BEGIN " "CHATFORMATTER DOTM-----"; e = "-----END " "CHATFORMATTER DOTM-----" }
        $0 == b { flag=1; next }
        $0 == e { flag=0; exit }
        flag { print }
    ' "$0" > "$TMP_B64"

    [ -s "$TMP_B64" ] || return 1
    decode_b64 "$TMP_B64" "$TMP_DOTM" || return 1

    local got_size got_sha
    got_size="$(wc -c < "$TMP_DOTM" | tr -d ' ')"
    [ "$got_size" = "$EXPECTED_SIZE" ] || return 1
    got_sha="$(calc_sha256 "$TMP_DOTM")" || return 1
    [ "$got_sha" = "$EXPECTED_SHA256" ] || return 1
    return 0
}

say "=============================================================="
say "  Chat Formatter for Word - Standalone Installer (macOS)"
say "=============================================================="
say ""

# --- 1. Extract and verify embedded package --------------------
say "[1/5] Preparing new ChatFormatter package..."
extract_payload || die "Could not extract or verify the embedded template."
say "      OK (verified $EXPECTED_SIZE bytes)"

# --- 2. Ensure Word is closed ----------------------------------
if [ "${CF_SKIP_WORD_CHECK:-0}" != "1" ] && command -v pgrep >/dev/null 2>&1 && pgrep -x "Microsoft Word" >/dev/null 2>&1; then
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
for dir in "${LEGACY_STARTUPS[@]}"; do
    if [ -d "$dir" ]; then
        rm -f "$dir"/ChatFormatter*.dotm "$dir"/~\$ChatFormatter*.dotm 2>/dev/null || true
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
printf 'AutoFormat = 1\nTables = 1\n' > "$SETTINGS" || die "Could not write $SETTINGS"
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

-----BEGIN CHATFORMATTER DOTM-----
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
-----END CHATFORMATTER DOTM-----
