#!/usr/bin/env bash
# One-shot install from WSL: install provider CLIs, build, trust the localhost
# certificate in Windows, start the bridge and sideload the add-in into Excel.
#
#   bash scripts/install-windows.sh                     # install / upgrade (default: copilot)
#   bash scripts/install-windows.sh --providers copilot,codex,opencode
#   bash scripts/install-windows.sh --reference-dir "/mnt/c/Users/me/Templates"
#   bash scripts/install-windows.sh --allowed-host 172.20.0.5   # extra hostname the bridge answers to (repeatable)
#   bash scripts/install-windows.sh --open              # ...and start a new Excel instance
#   bash scripts/install-windows.sh --no-providers      # skip provider install/login hints
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"

PROVIDERS=""
REFERENCE_DIR=""
ALLOWED_HOSTS=()
OPEN=0

usage() {
  sed -n '2,11p' "${BASH_SOURCE[0]}" | sed 's/^# \{0,1\}//'
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --providers) PROVIDERS="${2:-}"; shift 2 ;;
    --providers=*) PROVIDERS="${1#*=}"; shift ;;
    --reference-dir) REFERENCE_DIR="${2:-}"; shift 2 ;;
    --reference-dir=*) REFERENCE_DIR="${1#*=}"; shift ;;
    --allowed-host) ALLOWED_HOSTS+=("${2:-}"); shift 2 ;;
    --allowed-host=*) ALLOWED_HOSTS+=("${1#*=}"); shift ;;
    --open) OPEN=1; shift ;;
    --no-providers) PROVIDERS="none"; shift ;;
    -h|--help) usage; exit 0 ;;
    *) echo "unknown argument: $1" >&2; usage >&2; exit 2 ;;
  esac
done

command -v powershell.exe >/dev/null || { echo "powershell.exe not found: run this inside WSL2 with Windows interop enabled" >&2; exit 1; }

install_providers() {
  local spec="$1" entry name pkg bin login
  IFS=',' read -ra entries <<< "$spec"
  for entry in "${entries[@]}"; do
    name="$(echo "$entry" | tr -d ' ' | tr '[:upper:]' '[:lower:]')"
    case "$name" in
      copilot)  pkg="@github/copilot"; bin="copilot"; login="copilot login" ;;
      claude)   pkg="@anthropic-ai/claude-code"; bin="claude"; login="claude  (then /login)" ;;
      codex)    pkg="@openai/codex"; bin="codex"; login="codex login" ;;
      opencode) pkg="opencode-ai"; bin="opencode"; login="opencode auth login" ;;
      ""|none) continue ;;
      *) echo "    skipping unknown provider '$name'"; continue ;;
    esac
    if command -v "$bin" >/dev/null 2>&1; then
      echo "    $bin already installed"
    else
      echo "    installing $pkg (npm -g) ..."
      npm install -g "$pkg" || echo "    could not install $pkg; install it manually"
    fi
    echo "      login when ready:  $login"
  done
}

# Merge into the existing file with node (always present here) so other keys survive and no jq is needed.
write_config() {
  local ref="$1" file="$ROOT/excel-partner.config.json"
  [[ -z "$ref" && ${#ALLOWED_HOSTS[@]} -eq 0 ]] && return 0
  REF="$ref" HOSTS="${ALLOWED_HOSTS[*]:-}" CONFIG_FILE="$file" node -e '
    const fs = require("fs");
    let config = {};
    try { config = JSON.parse(fs.readFileSync(process.env.CONFIG_FILE, "utf8")); } catch {}
    if (process.env.REF) config.referenceDirs = [process.env.REF];
    const hosts = process.env.HOSTS.split(/\s+/).map(h => h.trim().toLowerCase()).filter(Boolean);
    if (hosts.length) config.allowedHosts = [...new Set([...(Array.isArray(config.allowedHosts) ? config.allowedHosts : []), ...hosts])];
    fs.writeFileSync(process.env.CONFIG_FILE, JSON.stringify(config, null, 2) + "\n");
  '
  [[ -n "$ref" ]] && echo "    wrote excel-partner.config.json (referenceDirs = $ref)"
  [[ ${#ALLOWED_HOSTS[@]} -gt 0 ]] && echo "    wrote excel-partner.config.json (allowedHosts += ${ALLOWED_HOSTS[*]})"
  return 0
}

if [[ -n "$PROVIDERS" ]]; then
  echo "==> provider CLIs"
  install_providers "$PROVIDERS"
fi

echo "==> build"
[[ -f certs/localhost.crt ]] || bash scripts/gen-cert.sh
node scripts/gen-icons.mjs >/dev/null
npm run build >/dev/null

echo "==> trust localhost certificate (current user)"
WIN_USERPROFILE="$(wslpath -u "$(powershell.exe -NoProfile -Command '$env:USERPROFILE' | tr -d '\r')")"
SHARE_WIN="$(wslpath -w "$WIN_USERPROFILE/ExcelAIAddin")"
mkdir -p "$WIN_USERPROFILE/ExcelAIAddin"
cp certs/localhost.crt "$WIN_USERPROFILE/ExcelAIAddin/localhost.crt"
if powershell.exe -NoProfile -Command '$t=(New-Object Security.Cryptography.X509Certificates.X509Certificate2 "'"$SHARE_WIN"'\localhost.crt").Thumbprint; if (Get-ChildItem Cert:\CurrentUser\Root | Where-Object Thumbprint -eq $t) { exit 0 } else { exit 1 }'; then
  echo "    already trusted"
else
  powershell.exe -NoProfile -Command 'certutil -addstore -user Root "'"$SHARE_WIN"'\localhost.crt"' | tail -1
fi

if [[ -n "$REFERENCE_DIR" || ${#ALLOWED_HOSTS[@]} -gt 0 ]]; then
  echo "==> local config"
  write_config "$REFERENCE_DIR"
fi

echo "==> (re)start bridge"
bash scripts/stop.sh >/dev/null 2>&1 || true
bash scripts/start.sh | tail -2

echo "==> sideload add-in"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$(wslpath -w scripts/sideload.ps1)" \
  -ManifestPath "$(wslpath -w addin/manifest.xml)" | grep -E "Copied|Registered" || true

if [[ "$OPEN" == "1" ]]; then
  echo "==> starting Excel"
  powershell.exe -NoProfile -Command 'Start-Process "C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE" -ArgumentList "/x"'
fi

cat <<MSG

Done. In Excel: Home > Excel Partner.
Restart any Excel window that was already open before this install.
MSG
