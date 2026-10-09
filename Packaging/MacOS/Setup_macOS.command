#!/bin/zsh
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
cd "$PROJECT_ROOT"

echo
echo "Setting up the macOS development environment..."
echo

if ! command -v python3 >/dev/null 2>&1; then
    echo "ERROR: python3 was not found. Install Python 3 for macOS, then run this file again."
    read -k 1 "?Press any key to close..."
    exit 1
fi

python3 -m venv .venv
".venv/bin/python" -m ensurepip --upgrade
".venv/bin/python" -m pip install --upgrade pip
".venv/bin/python" -m pip install -r requirements-macos.txt

mkdir -p "$HOME/Documents/Success Factors Weekly Training Report/Output"
mkdir -p "$HOME/Documents/Success Factors Weekly Training Report/Archive"
mkdir -p "$HOME/Documents/Success Factors Weekly Training Report/Logs"

echo
echo "SETUP COMPLETE"
echo "Run Packaging/macOS/Run_macOS.command to test the source version."
echo
read -k 1 "?Press any key to close..."
