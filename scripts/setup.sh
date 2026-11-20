#!/usr/bin/env bash
# pptmo (PPT Light) one-shot setup.
# Usage: setup.sh <target-dir> [--api-key KEY --api-base URL --model NAME]
# Copies app source, creates a venv, installs dependencies, optionally writes backend/.env.
set -euo pipefail

TARGET=""
API_KEY=""
API_BASE=""
MODEL=""

while [[ $# -gt 0 ]]; do
  case "$1" in
    --api-key) API_KEY="$2"; shift 2 ;;
    --api-base) API_BASE="$2"; shift 2 ;;
    --model) MODEL="$2"; shift 2 ;;
    *) TARGET="$1"; shift ;;
  esac
done

if [[ -z "$TARGET" ]]; then
  echo "usage: setup.sh <target-dir> [--api-key KEY --api-base URL --model NAME]" >&2
  exit 1
fi

SKILL_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
SRC="$SKILL_DIR/app"

if [[ ! -d "$SRC/backend" ]]; then
  echo "app source not found at $SRC (skill package incomplete?)" >&2
  exit 1
fi

command -v python3 >/dev/null || { echo "python3 not found" >&2; exit 1; }

if ! command -v soffice >/dev/null && ! command -v libreoffice >/dev/null; then
  echo "WARNING: soffice/libreoffice not on PATH — preview rendering will fail." >&2
  echo "         macOS: brew install --cask libreoffice ; Debian/Ubuntu: apt install libreoffice-impress" >&2
fi

mkdir -p "$TARGET"
cp -R "$SRC/." "$TARGET/"
cd "$TARGET/backend"

python3 -m venv .venv
.venv/bin/pip install --quiet --upgrade pip
.venv/bin/pip install --quiet -r requirements.txt

if [[ -n "$API_KEY" ]]; then
  umask 177
  {
    echo "DEEPSEEK_API_KEY=$API_KEY"
    [[ -n "$API_BASE" ]] && echo "DEEPSEEK_API_BASE=$API_BASE"
    [[ -n "$MODEL" ]] && echo "DEEPSEEK_MODEL=$MODEL"
    echo "RENDER_SCALE=2"
  } > .env
  echo "wrote backend/.env (600)"
else
  echo "no --api-key given: chat AI disabled; copy backend/.env.example to backend/.env to enable"
fi

echo
echo "setup complete. start with:"
echo "  cd $TARGET/backend && .venv/bin/python -m uvicorn main:app --host 127.0.0.1 --port 8090"
echo "then open http://127.0.0.1:8090"
