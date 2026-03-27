#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

if [ ! -d ".venv" ]; then
    echo "First-time setup required. Installing..."
    echo ""
    bash "$DIR/setup.sh"
    if [ $? -ne 0 ]; then
        echo ""
        echo "Setup failed. Press Enter to exit."
        read
        exit 1
    fi
    echo ""
fi

.venv/bin/python -m docx_redline.gui
echo ""
echo "Done. Press Enter to exit."
read
