#!/bin/bash
# ==============================================================
#   Chat Formatter for Word - macOS uninstaller
# ==============================================================
set -uo pipefail

STARTUP="$HOME/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
SETTINGS_DIR="$HOME/.chatformatter"

say() { printf '%s\n' "$*"; }
die() { printf '[ERROR] %s\n' "$*"; printf '\nPress any key to close...'; read -rsn1; exit 1; }

say "=============================================================="
say "  Chat Formatter for Word - Uninstall"
say "=============================================================="
say ""

if pgrep -x "Microsoft Word" >/dev/null 2>&1; then
    say "[1/2] Closing Word..."
    osascript -e 'tell application "Microsoft Word" to quit' >/dev/null 2>&1
    for _ in $(seq 1 20); do
        pgrep -x "Microsoft Word" >/dev/null 2>&1 || break
        sleep 0.5
    done
else
    say "[1/2] Word is not running."
fi

say "[2/2] Removing the add-in..."
rm -f "$STARTUP/ChatFormatter.dotm" && say "      Removed the template." || say "      Nothing to remove."

if [ -d "$SETTINGS_DIR" ]; then
    rm -rf "$SETTINGS_DIR" && say "      Removed your settings." || say "      Could not remove your settings."
else
    say "      No settings to remove."
fi

say ""
say "Done. Word no longer has Chat Formatter loaded."
say ""
printf 'Press any key to close...'
read -rsn1
printf '\n'
