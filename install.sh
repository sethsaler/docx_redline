#!/bin/bash
set -e

REPO="sethsaler/docx_redline"
BRANCH="main"
DEST="$HOME/Desktop/docx_redline"

RED='\033[0;31m'
GREEN='\033[0;32m'
BOLD='\033[1m'
YELLOW='\033[0;33m'
CYAN='\033[0;36m'
RESET='\033[0m'

echo ""
echo "========================================"
echo "  DOCX Redline — Installer"
echo "========================================"
echo ""

if command -v git &>/dev/null; then
    echo "Downloading via git..."
    git clone --depth 1 "https://github.com/${REPO}.git" "$DEST" 2>/dev/null || {
        echo -e "${RED}Download failed. Check your internet connection.${RESET}"
        exit 1
    }
    rm -rf "$DEST/.git"
else
    echo "Downloading..."
    TMPFILE=$(mktemp /tmp/docx_redline.XXXXXX.zip)
    URL="https://github.com/${REPO}/archive/refs/heads/${BRANCH}.zip"

    if ! curl -sL -o "$TMPFILE" "$URL"; then
        echo -e "${RED}Download failed. Check your internet connection.${RESET}"
        rm -f "$TMPFILE"
        exit 1
    fi

    unzip -q "$TMPFILE" -d /tmp/docx_redline_extract 2>/dev/null
    mv "/tmp/docx_redline_extract/docx_redline-${BRANCH}" "$DEST"
    rm -f "$TMPFILE"
    rm -rf /tmp/docx_redline_extract
fi

echo -e "${GREEN}Downloaded.${RESET}"
echo ""

cd "$DEST"
bash setup.sh
