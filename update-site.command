#!/bin/bash
# Double-click this file to apply your edits from the updates/ folder.
# It works on macOS. The Windows equivalent is update-site.bat.

set -e
cd "$(dirname "$0")"

echo "================================================"
echo "  MUSE Lab — Update Site"
echo "================================================"
echo ""

# First-run setup: create venv and install dependencies if missing.
if [ ! -d ".venv" ]; then
    echo "First-time setup: creating a Python virtual environment..."
    echo "(this only happens once and takes about 30 seconds)"
    echo ""
    if ! command -v python3 >/dev/null 2>&1; then
        echo "ERROR: python3 is not installed."
        echo "Install Python 3 from https://www.python.org/downloads/"
        echo ""
        read -n 1 -s -r -p "Press any key to close..."
        exit 1
    fi
    python3 -m venv .venv
    .venv/bin/pip install --quiet --upgrade pip
    .venv/bin/pip install --quiet -r scripts/requirements.txt
    echo "Setup complete."
    echo ""
fi

echo "Reading your Excel sheets and applying changes..."
echo ""
.venv/bin/python scripts/update_site.py

echo ""
echo "================================================"
echo "  Done."
echo ""
echo "  Next steps:"
echo "    1. Open this folder in GitHub Desktop (or your"
echo "       favorite git tool) to see what changed."
echo "    2. Commit and push to publish your changes."
echo "================================================"
echo ""
read -n 1 -s -r -p "Press any key to close this window..."
echo ""
