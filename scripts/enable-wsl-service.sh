#!/usr/bin/env bash
# Run the bridge as a persistent WSL user service, independent of the shell that starts it.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SERVICE="$ROOT/run/excel-partner-bridge.service"
COPILOT_CLI="$(npm prefix -g)/bin/copilot"
if [[ ! -x "$COPILOT_CLI" ]]; then
  echo "Copilot CLI not found at $COPILOT_CLI" >&2
  exit 1
fi
mkdir -p "$ROOT/run"

cat > "$SERVICE" <<EOF
[Unit]
Description=Excel Partner bridge
After=network.target

[Service]
Type=simple
WorkingDirectory=$ROOT
Environment=COPILOT_BIN=$COPILOT_CLI
ExecStart=/usr/bin/bash -lc 'exec /usr/bin/node "$ROOT/dist/bridge/src/index.js"'
Restart=on-failure
RestartSec=2

[Install]
WantedBy=default.target
EOF

systemctl --user link "$SERVICE"
systemctl --user daemon-reload
systemctl --user enable --now excel-partner-bridge.service
systemctl --user restart excel-partner-bridge.service
systemctl --user is-active --quiet excel-partner-bridge.service
echo "WSL bridge service active: excel-partner-bridge.service"
