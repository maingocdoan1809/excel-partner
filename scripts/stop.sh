#!/usr/bin/env bash
# Stop the background bridge started by scripts/start.sh.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PID_FILE="$ROOT/run/bridge.pid"

if command -v systemctl >/dev/null 2>&1 && systemctl --user is-enabled --quiet excel-partner-bridge.service 2>/dev/null; then
  systemctl --user stop excel-partner-bridge.service
  rm -f "$PID_FILE"
  echo "bridge stopped (excel-partner-bridge.service)"
  exit 0
fi

if [[ ! -f "$PID_FILE" ]]; then
  echo "bridge is not running (no pid file)"
  exit 0
fi

PID="$(cat "$PID_FILE")"
if kill -0 "$PID" 2>/dev/null; then
  kill "$PID"
  for _ in $(seq 1 20); do
    kill -0 "$PID" 2>/dev/null || break
    sleep 0.25
  done
  kill -0 "$PID" 2>/dev/null && kill -9 "$PID"
  echo "bridge stopped (pid $PID)"
else
  echo "bridge was not running (stale pid $PID)"
fi
rm -f "$PID_FILE"
