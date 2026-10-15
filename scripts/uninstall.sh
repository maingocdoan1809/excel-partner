#!/usr/bin/env bash
# Stop the bridge and remove generated artifacts. Source files are kept.
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

echo "==> stopping bridge"
bash scripts/stop.sh || true

echo "==> removing build output"
rm -rf "$ROOT/dist" "$ROOT/addin/dist" "$ROOT/run"

echo "==> removing generated certificate"
rm -rf "$ROOT/certs"

echo
echo "Uninstalled local artifacts."
echo "Remember to also remove the 'Excel Partner' add-in from Excel"
echo "(Insert > My Add-ins > right-click > Remove) and delete the manifest"
echo "from the shared folder used by scripts/sideload.ps1."
