#!/bin/bash
# ==============================================================
#   Chat Formatter for Word - macOS installer
#   Copies the template into Word's STARTUP folder and writes the
#   default settings.  No registry, no helper, no admin rights.
# ==============================================================
set -uo pipefail

DIR="$(cd "$(dirname "$0")" && pwd)"
TEMPLATE="ChatFormatter.dotm"
STARTUP="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
SETTINGS_DIR="$HOME/.chatformatter"
SETTINGS="$SETTINGS_DIR/settings.ini"

say() { printf '%s\n' "$*"; }
die() { printf '[ERROR] %s\n' "$*"; printf '\nPress any key to close...'; read -rsn1; exit 1; }

[ -f "$DIR/$TEMPLATE" ] || die "$TEMPLATE not found next to this script."

say "=============================================================="
say "  Chat Formatter for Word - Install"
say "=============================================================="
say ""

# --- close Word so the add-in is not half-loaded ----------------
if pgrep -x "Microsoft Word" >/dev/null 2>&1; then
    say "[1/4] Word is running. Closing it..."
    osascript -e 'tell application "Microsoft Word" to quit' >/dev/null 2>&1
    for _ in $(seq 1 20); do
        pgrep -x "Microsoft Word" >/dev/null 2>&1 || break
        sleep 0.5
    done
    pgrep -x "Microsoft Word" >/dev/null 2>&1 && die "Word did not close. Quit it and run this again."
else
    say "[1/4] Word is not running."
fi

# --- copy the template -----------------------------------------
say "[2/4] Installing into the Word STARTUP folder..."
mkdir -p "$STARTUP" || die "Could not create $STARTUP"
if ! cp "$DIR/$TEMPLATE" "$STARTUP/$TEMPLATE"; then
    say "      Copy failed, retrying after removing the old copy."
    rm -f "$STARTUP/$TEMPLATE"
    cp "$DIR/$TEMPLATE" "$STARTUP/$TEMPLATE" || die "Copy failed. Check Disk Access for your user account."
fi
say "      OK -> $STARTUP/$TEMPLATE"

# --- default settings ------------------------------------------
say "[3/4] Writing default settings..."
mkdir -p "$SETTINGS_DIR"
if [ ! -f "$SETTINGS" ]; then
    printf 'AutoFormat = 1\nTables = 1\n' > "$SETTINGS"
    say "      Created $SETTINGS"
else
    say "      Kept your existing $SETTINGS"
fi

# --- first run needs a manual trust decision -------------------
say "[4/4] Macro security"
say "      The first time Word opens the template it asks whether to"
say "      enable the macros. Choose Yes. After that it loads silently."
say ""
say "      If the prompt does not appear: Word > Settings > Security"
say "      (or Preferences > Security) > Trusted Locations."
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
say "  Turn tables off?              Tools > Macro > ToggleTables"
say ""
printf 'Press any key to close...'
read -rsn1
printf '\n'
