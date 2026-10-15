#!/usr/bin/env bash
# First-run setup: install dependencies, build, generate the TLS certificate.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

echo "==> node version"
node --version

echo "==> installing dependencies"
npm install

echo "==> generating localhost certificate"
bash scripts/gen-cert.sh

echo "==> building bridge, mcp server and task pane"
npm run build

echo
echo "Setup complete. Next:"
echo "  1) bash scripts/start.sh"
echo "  2) powershell -File scripts/sideload.ps1   (from Windows)"
echo "  3) open Excel and load the 'Excel Partner' add-in"
