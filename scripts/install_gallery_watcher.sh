#!/bin/zsh
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd -P)"
LABEL="com.nolimit.gallery-watcher"
PLIST="$HOME/Library/LaunchAgents/${LABEL}.plist"
PYTHON_BIN="$(command -v python3)"
WATCHER="$ROOT/scripts/gallery_watch.py"
LOG_DIR="$ROOT/scripts/logs"

case "$ROOT" in
  "$HOME/Desktop"/*|"$HOME/Documents"/*|"$HOME/Downloads"/*)
    echo "Nao foi possivel instalar o watcher em segundo plano a partir desta pasta."
    echo "Motivo: o macOS costuma bloquear LaunchAgents acessando projetos dentro de Desktop, Documents ou Downloads."
    echo "Mova o projeto para uma pasta como $HOME/Projects e execute novamente este instalador."
    exit 1
    ;;
esac

mkdir -p "$HOME/Library/LaunchAgents"
mkdir -p "$LOG_DIR"

cat > "$PLIST" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${LABEL}</string>
  <key>ProgramArguments</key>
  <array>
    <string>${PYTHON_BIN}</string>
    <string>${WATCHER}</string>
PLIST

for arg in "$@"; do
  printf '    <string>%s</string>\n' "$arg" >> "$PLIST"
done

cat >> "$PLIST" <<PLIST
  </array>
  <key>KeepAlive</key>
  <true/>
  <key>RunAtLoad</key>
  <true/>
  <key>WorkingDirectory</key>
  <string>${ROOT}</string>
  <key>StandardOutPath</key>
  <string>${LOG_DIR}/gallery-watch-launchd.out.log</string>
  <key>StandardErrorPath</key>
  <string>${LOG_DIR}/gallery-watch-launchd.err.log</string>
</dict>
</plist>
PLIST

launchctl bootout "gui/$(id -u)" "$PLIST" >/dev/null 2>&1 || true
launchctl bootstrap "gui/$(id -u)" "$PLIST"
launchctl kickstart -k "gui/$(id -u)/${LABEL}"

echo "Gallery watcher instalado e iniciado."
echo "Plist: $PLIST"
echo "Watcher log: $LOG_DIR/gallery-watch.log"
echo "Launchd stdout: $LOG_DIR/gallery-watch-launchd.out.log"
echo "Launchd stderr: $LOG_DIR/gallery-watch-launchd.err.log"
