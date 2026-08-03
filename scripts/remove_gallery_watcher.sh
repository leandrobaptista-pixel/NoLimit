#!/bin/zsh
set -euo pipefail

LABEL="com.nolimit.gallery-watcher"
PLIST="$HOME/Library/LaunchAgents/${LABEL}.plist"

launchctl bootout "gui/$(id -u)" "$PLIST" >/dev/null 2>&1 || true
rm -f "$PLIST"

echo "Gallery watcher removido."
echo "Se quiser, os logs antigos continuam em scripts/logs/."
