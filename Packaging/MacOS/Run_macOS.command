#!/bin/zsh
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
cd "$PROJECT_ROOT"

if [[ ! -x ".venv/bin/python" ]]; then
    echo "ERROR: .venv/bin/python was not found."
    echo "Run Packaging/macOS/Setup_macOS.command first."
    read -k 1 "?Press any key to close..."
    exit 1
fi

".venv/bin/python" "Scripts/gsf_app.py"
