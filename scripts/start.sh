#!/usr/bin/env bash
# Start the bridge in the background, recording its pid and log.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

RUN_DIR="$ROOT/run"
PID_FILE="$RUN_DIR/bridge.pid"
LOG_FILE="$RUN_DIR/bridge.log"
mkdir -p "$RUN_DIR"

if command -v systemctl >/dev/null 2>&1 && systemctl --user is-enabled --quiet excel-partner-bridge.service 2>/dev/null; then
  systemctl --user start excel-partner-bridge.service
  systemctl --user is-active --quiet excel-partner-bridge.service
  echo "bridge running via excel-partner-bridge.service"
  exit 0
fi

if [[ -f "$PID_FILE" ]] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
  echo "bridge already running (pid $(cat "$PID_FILE"))"
  exit 0
fi

if [[ ! -f "$ROOT/dist/bridge/src/index.js" ]]; then
  echo "build output missing; running 'npm run build'"
  npm run build
fi

echo "starting bridge..."
nohup node "$ROOT/dist/bridge/src/index.js" >>"$LOG_FILE" 2>&1 &
echo $! > "$PID_FILE"

sleep 1
if kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
  echo "bridge started (pid $(cat "$PID_FILE"))"
  echo "log: $LOG_FILE"
else
  echo "bridge failed to start; see $LOG_FILE"
  tail -n 30 "$LOG_FILE" || true
  exit 1
fi
