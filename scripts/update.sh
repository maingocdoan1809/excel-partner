#!/usr/bin/env bash
# Update an existing install: pull, reinstall, rebuild, re-sideload, restart.
#
#   bash scripts/update.sh            # update
#   bash scripts/update.sh --open     # ...and start a new Excel instance
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

OPEN=0
while [[ $# -gt 0 ]]; do
  case "$1" in
    --open) OPEN=1; shift ;;
    -h|--help) sed -n '2,5p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'; exit 0 ;;
    *) echo "unknown argument: $1" >&2; exit 2 ;;
  esac
done

if [[ -d .git ]]; then
  echo "==> pulling latest changes"
  git pull --ff-only || echo "    git pull failed; continuing with the current checkout"
fi

echo "==> installing dependencies"
if [[ -f package-lock.json ]]; then npm ci; else npm install; fi

echo "==> rebuilding"
npm run build

echo "==> restarting bridge"
bash scripts/stop.sh || true
bash scripts/start.sh

echo "==> re-sideloading add-in"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(wslpath -w scripts/sideload.ps1)" \
  -ManifestPath "$(wslpath -w addin/manifest.xml)" | grep -E "Copied|Registered" || true

if [[ "$OPEN" == "1" ]]; then
  echo "==> starting Excel"
  powershell.exe -NoProfile -Command 'Start-Process "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" -ArgumentList "/x"'
fi

echo "Update complete."
