#!/bin/bash
set -e

DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

REPO="sethsaler/docx_redline"
BRANCH="main"

RED='\033[0;31m'
GREEN='\033[0;32m'
BOLD='\033[1m'
YELLOW='\033[0;33m'
CYAN='\033[0;36m'
RESET='\033[0m'

echo ""
echo "========================================"
echo "  DOCX Redline — Update"
echo "========================================"
echo ""

update_via_zip() {
    echo "Downloading latest ZIP from GitHub..."
    local TMPFILE
    TMPFILE=$(mktemp /tmp/docx_redline.XXXXXX.zip)
    local URL="https://github.com/${REPO}/archive/refs/heads/${BRANCH}.zip"

    if ! curl -fsSL -o "$TMPFILE" "$URL"; then
        echo -e "${RED}Download failed. Check your internet connection or try again later.${RESET}"
        rm -f "$TMPFILE"
        return 1
    fi

    rm -rf /tmp/docx_redline_update_extract
    mkdir -p /tmp/docx_redline_update_extract
    if ! unzip -q "$TMPFILE" -d /tmp/docx_redline_update_extract; then
        echo -e "${RED}Could not unpack the download.${RESET}"
        rm -f "$TMPFILE"
        rm -rf /tmp/docx_redline_update_extract
        return 1
    fi
    rm -f "$TMPFILE"

    local SRC="/tmp/docx_redline_update_extract/docx_redline-${BRANCH}"
    if [ ! -d "$SRC" ]; then
        echo -e "${RED}Unexpected archive layout.${RESET}"
        rm -rf /tmp/docx_redline_update_extract
        return 1
    fi

    echo "Merging files into this folder (keeping your .venv)..."
    rsync -a --delete \
        --exclude='.venv' \
        --exclude='.git' \
        "${SRC}/" "$DIR/"
    rm -rf /tmp/docx_redline_update_extract
    echo -e "${GREEN}Update files applied.${RESET}"
    return 0
}

if [ -d "$DIR/.git" ] && command -v git &>/dev/null; then
    echo "Updating via git..."
    if ! git -C "$DIR" rev-parse --git-dir &>/dev/null; then
        echo -e "${YELLOW}Not a valid git repository; falling back to ZIP download...${RESET}"
        echo ""
        if ! update_via_zip; then
            exit 1
        fi
    else
        git -C "$DIR" fetch origin "$BRANCH" 2>/dev/null || true
        if ! git -C "$DIR" pull --ff-only "origin" "$BRANCH"; then
            echo ""
            echo -e "${RED}Git pull failed.${RESET} If you changed files here, stash or commit them, then try again."
            echo -e "Or remove the ${CYAN}.git${RESET} folder and run this script again to use ZIP updates instead."
            exit 1
        fi
        echo -e "${GREEN}Repository updated.${RESET}"
    fi
else
    if ! update_via_zip; then
        exit 1
    fi
fi

echo ""
echo "Refreshing Python environment and package..."
bash "$DIR/setup.sh"

echo ""
echo -e "${BOLD}Update complete.${RESET}"
echo ""
