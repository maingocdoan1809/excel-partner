#!/usr/bin/env bash
# Install the working copy (WSL) as the live Office Partner on this machine's Windows side.
#   npm run deploy:windows
# Packs the release tarball, copies it to the Windows temp folder and runs the installed `office-partner.cmd update --from` there.
# The updater must run on Windows with the installed CLI: started from a \\wsl.localhost path it refuses (not a global install)
# and can leave a gateway running from the WSL folder. Close/reopen Word, Excel and PowerPoint afterwards.
set -euo pipefail
cd "$(dirname "$0")/.."

CURL=/mnt/c/Windows/System32/curl.exe
WIN_TEMP=$(wslpath "$(cmd.exe /c "echo %TEMP%" 2>/dev/null | tr -d '\r')")
WORK="$WIN_TEMP/op-deploy"
VERSION=$(node -p "require('./package.json').version")

health=$($CURL -sk -m 5 https://localhost:8765/health || true)
if [[ "$health" == *'"busy":true'* ]]; then echo "A chat is running in Office; wait for it to finish, then run again." >&2; exit 1; fi

npm run pack:release >/dev/null
mkdir -p "$WORK"
cp "release/office-partner-$VERSION.tgz" "$WORK/"
rm -f "$WORK/update.txt"
cat > "$WORK/update.ps1" <<PS
Set-Location \$env:TEMP
\$out = Join-Path \$env:TEMP "op-deploy\\update.txt"
\$tgz = Join-Path \$env:TEMP "op-deploy\\office-partner-$VERSION.tgz"
& (Join-Path \$env:APPDATA "npm\\office-partner.cmd") update --from \$tgz *>&1 | Out-String | Out-File -Encoding ascii \$out
"EXIT " + \$LASTEXITCODE | Out-File -Append -Encoding ascii \$out
PS

echo "Updating Windows to $VERSION (1-3 minutes)..."
# The services the update restarts keep powershell's pipes open, so wait for the result file instead of for the process.
powershell.exe -ExecutionPolicy Bypass -File "$(wslpath -w "$WORK/update.ps1")" >/dev/null 2>&1 </dev/null &
for _ in $(seq 1 240); do grep -aq "^EXIT" "$WORK/update.txt" 2>/dev/null && break; sleep 3; done
grep -aq "^EXIT" "$WORK/update.txt" 2>/dev/null || { echo "The update did not finish in 12 minutes; see $WORK/update.txt" >&2; exit 1; }
# The update's output may contain the gateway token; show only the progress lines.
grep -aE "^(Update|Bridge|Registered|Installing|EXIT)|Error|failed" "$WORK/update.txt" | grep -av "[Tt]oken" || true

$CURL -sk -m 5 https://localhost:8765/health | grep -q '"status":"ok"' && echo "Bridge: ok" || { echo "Bridge is not answering." >&2; exit 1; }
$CURL -sk -m 5 https://localhost:8766/health | grep -q '"ready":true' && echo "Gateway: ok" || echo "Gateway is not healthy; run: office-partner gateway start" >&2
echo "Reopen Word, Excel and PowerPoint to load the new pane."
