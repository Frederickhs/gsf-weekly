#!/bin/zsh
set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
cd "$PROJECT_ROOT"

APP_NAME="Success Factors Weekly Training Report"
APP_PATH="dist/$APP_NAME.app"
BUILD_ARCHIVE="Archive/Application Builds/macOS"
SPEC_FILE="Packaging/macOS/SuccessFactorsWeeklyTrainingReport-macOS.spec"

echo
echo "Building $APP_NAME for macOS..."
echo

if [[ ! -x ".venv/bin/python" ]]; then
    echo "ERROR: .venv/bin/python was not found."
    echo "Run Packaging/macOS/Setup_macOS.command first."
    read -k 1 "?Press any key to close..."
    exit 1
fi

if [[ -d "$APP_PATH" ]]; then
    TIMESTAMP="$(date +%Y-%m-%d_%H%M%S)"
    mkdir -p "$BUILD_ARCHIVE"
    echo "Archiving previous macOS application build..."
    mv "$APP_PATH" "$BUILD_ARCHIVE/${APP_NAME}_${TIMESTAMP}.app"
fi

".venv/bin/python" -m PyInstaller --noconfirm --clean "$SPEC_FILE"

cp -f VERSION.txt "dist/VERSION.txt"
cp -f README.txt "dist/README.txt"

echo
echo "BUILD COMPLETE"
echo "App: $PROJECT_ROOT/$APP_PATH"
echo "Writable reports: ~/Documents/$APP_NAME"
echo
open "$PROJECT_ROOT/dist"
read -k 1 "?Press any key to close..."
