#!/usr/bin/env bash
# Install "Excel Partner" from scratch (WSL2 path).
#
# Run from anywhere, even before cloning. Reuses the current checkout if this
# file sits inside one, otherwise clones the repo into $REPO_DIR, then hands
# over to scripts/install-windows.sh.
#
#   bash install.sh                          # clone (if needed) + install
#   bash install.sh --providers copilot,codex,opencode --open
#   REPO_DIR=~/tools/excel-partner bash install.sh
set -euo pipefail

REPO_URL="${REPO_URL:-}"
REPO_SLUG="${REPO_SLUG:-}"
REPO_DIR="${REPO_DIR:-$HOME/excel-partner}"

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

if [[ -f "$SCRIPT_DIR/package.json" && -f "$SCRIPT_DIR/scripts/install-windows.sh" ]]; then
  ROOT="$SCRIPT_DIR"
else
  echo "==> bootstrapping: cloning $REPO_SLUG"
  if [[ -d "$REPO_DIR/.git" ]]; then
    echo "    using existing checkout at $REPO_DIR"
    git -C "$REPO_DIR" pull --ff-only || echo "    git pull failed; continuing"
  elif [[ -e "$REPO_DIR" && -n "$(ls -A "$REPO_DIR" 2>/dev/null)" ]]; then
    echo "ERROR: $REPO_DIR exists and is not a git checkout." >&2
    exit 1
  elif [[ -z "$REPO_URL" && -z "$REPO_SLUG" ]]; then
    echo "ERROR: set REPO_URL or REPO_SLUG when installing outside a checkout." >&2
    exit 1
  elif command -v gh >/dev/null 2>&1; then
    gh repo clone "${REPO_SLUG:-$REPO_URL}" "$REPO_DIR" || { echo "gh clone failed (run 'gh auth login')" >&2; exit 1; }
  elif command -v git >/dev/null 2>&1; then
    git clone "${REPO_URL:-https://github.com/$REPO_SLUG.git}" "$REPO_DIR" || { echo "git clone failed (private repo needs credentials)" >&2; exit 1; }
  else
    echo "ERROR: neither gh nor git found. Install git first." >&2
    exit 1
  fi
  ROOT="$REPO_DIR"
fi

exec bash "$ROOT/scripts/install-windows.sh" "$@"
