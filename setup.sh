#!/bin/bash
set -e

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

RED='\033[0;31m'
GREEN='\033[0;32m'
BOLD='\033[1m'
YELLOW='\033[0;33m'
CYAN='\033[0;36m'
RESET='\033[0m'

FALLBACK_PYTHON_VERSION="3.13.2"

echo ""
echo "========================================"
echo "  DOCX Redline — Setup"
echo "========================================"
echo ""

find_python() {
    local candidates=(
        "/opt/homebrew/bin/python3.14"
        "/opt/homebrew/bin/python3.13"
        "/opt/homebrew/bin/python3.12"
        "/opt/homebrew/bin/python3.11"
        "/opt/homebrew/bin/python3.10"
        "/opt/homebrew/bin/python3.9"
        "/opt/homebrew/bin/python3"
        "/usr/local/bin/python3"
        "/Library/Frameworks/Python.framework/Versions/3.14/bin/python3"
        "/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
        "/Library/Frameworks/Python.framework/Versions/3.12/bin/python3"
        "/Library/Frameworks/Python.framework/Versions/3.11/bin/python3"
        "/LibraryFrameworks/Python.framework/Versions/3.10/bin/python3"
        "/Library/Frameworks/Python.framework/Versions/3.9/bin/python3"
    )

    for cmd in python3.14 python3.13 python3.12 python3.11 python3.10 python3.9 python3 python; do
        if command -v "$cmd" &>/dev/null; then
            candidates+=("$(command -v "$cmd")")
        fi
    done

    local best=""
    local best_v=0

    for cmd in "${candidates[@]}"; do
        if [ -x "$cmd" ] 2>/dev/null; then
            local ver
            ver=$("$cmd" -c "import sys; v=sys.version_info; print(v.major*100+v.minor)" 2>/dev/null) || continue
            if [ "$ver" -ge 309 ] && [ "$ver" -gt "$best_v" ]; then
                best="$cmd"
                best_v="$ver"
            fi
        fi
    done

    echo "$best"
}

install_python() {
    echo -e "${YELLOW}Python 3.9+ not found on this Mac.${RESET}"
    echo ""
    echo "Attempting to download and install Python..."
    echo ""

    local latest="$FALLBACK_PYTHON_VERSION"
    local detected
    detected=$(curl -s https://www.python.org/ftp/python/ 2>/dev/null \
        | grep -oE 'href="3\.[0-9]+\.[0-9]+/"' \
        | grep -oE '3\.[0-9]+\.[0-9]+' \
        | sort -t. -k1,1n -k2,2n -k3,3n \
        | tail -1)

    if [ -n "$detected" ]; then
        latest="$detected"
    fi

    local url="https://www.python.org/ftp/python/${latest}/python-${latest}-macos11.pkg"
    local pkg="/tmp/python-${latest}-macos11.pkg"

    echo "  Downloading Python ${latest}..."
    if ! curl -L -o "$pkg" "$url" --progress-bar --fail 2>/dev/null; then
        echo ""
        echo -e "${RED}Download failed.${RESET}"
        echo ""
        echo "Please install Python manually:"
        echo -e "  ${CYAN}https://www.python.org/downloads/${RESET}"
        echo ""
        echo "Then re-run this script."
        exit 1
    fi

    echo ""
    echo "  Installing Python ${latest}..."
    echo -e "  ${YELLOW}(You may be prompted for your administrator password.)${RESET}"
    echo ""

    if sudo installer -pkg "$pkg" -target / 2>/dev/null; then
        rm -f "$pkg"
        echo -e "${GREEN}  Python ${latest} installed.${RESET}"
        echo ""
    else
        rm -f "$pkg"
        echo ""
        echo -e "${YELLOW}Command-line install was canceled or failed.${RESET}"
        echo "Opening the graphical installer instead..."
        echo ""
        open "$url"
        echo -e "${YELLOW}Please complete the Python installation, then re-run this script.${RESET}"
        echo ""
        exit 1
    fi
}

# --- Step 1: Find or install Python ---

PYTHON_CMD=$(find_python)

if [ -z "$PYTHON_CMD" ]; then
    install_python

    PYTHON_CMD=$(find_python)

    if [ -z "$PYTHON_CMD" ]; then
        echo -e "${RED}Python was installed but could not be located.${RESET}"
        echo "Please open a new Terminal window and re-run this script."
        exit 1
    fi
fi

PY_VER=$("$PYTHON_CMD" -c "import sys; v=sys.version_info; print(f'{v.major}.{v.minor}.{v.micro}')")
echo -e "  Found Python ${PY_VER}  (${PYTHON_CMD})"

# --- Step 2: Create virtual environment ---

if [ -d ".venv" ]; then
    echo -e "${GREEN}  Virtual environment already exists.${RESET} Reusing it."
else
    echo "  Creating virtual environment..."
    "$PYTHON_CMD" -m venv .venv
    echo -e "  ${GREEN}Done.${RESET}"
fi

# --- Step 3: Install the tool ---

echo "  Installing docx-redline and dependencies..."
.venv/bin/pip install --upgrade pip --quiet 2>/dev/null
.venv/bin/pip install . --quiet 2>/dev/null
echo -e "  ${GREEN}Done.${RESET}"

# Tk (tkinter) is required for the graphical file picker in run.command.
if ! .venv/bin/python -c "import tkinter" 2>/dev/null; then
    echo ""
    echo -e "${YELLOW}  Note: Tk (tkinter) is not available in this Python.${RESET}"
    echo -e "  The window with Browse buttons will not open; the tool will use terminal prompts instead."
    echo ""
    if [[ "$(uname -s)" == "Darwin" ]]; then
        echo -e "  On macOS, install a Python build that includes Tcl/Tk, for example:"
        echo -e "    ${CYAN}https://www.python.org/downloads/${RESET} (official installer — include Tcl/Tk)"
        echo -e "  Or with Homebrew: ${CYAN}brew install python-tk${RESET} then point setup at that Python."
    else
        echo -e "  On Linux, install your distro's Tk package, e.g.: ${CYAN}sudo apt install python3-tk${RESET}"
    fi
    echo ""
fi

echo ""
echo -e "${BOLD}Setup complete!${RESET}"
echo ""
echo "To use the tool:"
echo "  Double-click ${CYAN}run.command${RESET}"
echo "  Or run: ${CYAN}./run.command${RESET}"
echo ""
