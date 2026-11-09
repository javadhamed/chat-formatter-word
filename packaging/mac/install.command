#!/bin/bash
# ==============================================================
#   Chat Formatter for Word - macOS installer
#   Automatically removes any previous version first, then copies
#   the template into Word's STARTUP folder and writes fresh
#   default settings.  No registry, no helper, no admin rights.
# ==============================================================
set -uo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
TEMPLATE="ChatFormatter.dotm"
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

[ -f "$DIR/$TEMPLATE" ] || die "$TEMPLATE not found next to this script. Use ChatFormatter-Mac-Installer.command for standalone install."

say "=============================================================="
say "  Chat Formatter for Word - Install"
say "=============================================================="
say ""

# --- close Word so the add-in is not half-loaded ----------------
if [ "${CF_SKIP_WORD_CHECK:-0}" != "1" ] && command -v pgrep >/dev/null 2>&1 && pgrep -x "Microsoft Word" >/dev/null 2>&1; then
    say "[1/4] Word is running. Closing it..."
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
else
    say "[1/4] Word is not running."
fi

# --- remove previous version completely -------------------------
say "[2/4] Removing any previous version..."
for dir in "${LEGACY_STARTUPS[@]}"; do
    if [ -d "$dir" ]; then
        rm -f "$dir"/ChatFormatter*.dotm "$dir"/~\$ChatFormatter*.dotm 2>/dev/null || true
    fi
done
if [ -d "$SETTINGS_DIR" ]; then
    rm -rf "$SETTINGS_DIR" || die "Could not remove old settings at $SETTINGS_DIR"
fi
say "      Previous files and settings cleaned."

# --- copy the template ------------------------------------------
say "[3/4] Installing into the Word STARTUP folder..."
mkdir -p "$STARTUP" || die "Could not create $STARTUP"
cp "$DIR/$TEMPLATE" "$STARTUP/$TEMPLATE" || die "Copy failed. Check Disk Access for your user account."
chmod 644 "$STARTUP/$TEMPLATE" 2>/dev/null || true
if command -v xattr >/dev/null 2>&1; then
    xattr -d com.apple.quarantine "$STARTUP/$TEMPLATE" 2>/dev/null || true
fi
say "      OK -> $STARTUP/$TEMPLATE"

# --- default settings -------------------------------------------
say "[4/4] Writing fresh default settings..."
mkdir -p "$SETTINGS_DIR" || die "Could not create $SETTINGS_DIR"
printf 'AutoFormat = 1\nTables = 1\n' > "$SETTINGS" || die "Could not write $SETTINGS"
say "      Created $SETTINGS"

say ""
say "      First run: Word asks whether to enable the macros — choose Yes."
say "      (Word > Settings > Security > Trusted Locations if needed.)"
say ""
say "=============================================================="
say "  Installation complete."
say "=============================================================="
say ""
say "  How to use:"
say "    1. Copy an answer from ChatGPT, DeepSeek, Gemini, ..."
say "    2. Paste it into Word"
say "    3. Select the pasted text and RIGHT-CLICK"
say ""
say "  Prefer an explicit button?  Tools > Macro > FormatChatText"
say "  Turn the automatic mode off?  Tools > Macro > ToggleAutoFormat"

pause_exit 0
