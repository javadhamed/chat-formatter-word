#!/bin/bash
# ==============================================================
#   Chat Formatter for Word - Standalone macOS Installer
#   Self-contained: embeds ChatFormatter.dotm, removes any old
#   installation first, then installs the new version.
# ==============================================================
set -uo pipefail

TEMPLATE="ChatFormatter.dotm"
EXPECTED_SIZE=30788
EXPECTED_SHA256="63fbb9322b4a14c82ea9d7de44bf4a6772f161b6923b734f273e2b773dabcf08"

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
AACidAAAAAA=
-----END CHATFORMATTER DOTM-----
