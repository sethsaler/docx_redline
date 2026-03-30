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

download_via_zip() {
    echo "Downloading ZIP archive from GitHub..."
    local TMPFILE
    TMPFILE=$(mktemp /tmp/docx_redline.XXXXXX.zip)
    local URL="https://github.com/${REPO}/archive/refs/heads/${BRANCH}.zip"

    if ! curl -fsSL -o "$TMPFILE" "$URL"; then
        echo -e "${RED}Download failed. Check your internet connection or try again later.${RESET}"
        rm -f "$TMPFILE"
        return 1
    fi

    rm -rf /tmp/docx_redline_extract
    mkdir -p /tmp/docx_redline_extract
    if ! unzip -q "$TMPFILE" -d /tmp/docx_redline_extract; then
        echo -e "${RED}Could not unpack the download.${RESET}"
        rm -f "$TMPFILE"
        rm -rf /tmp/docx_redline_extract
        return 1
    fi
    rm -f "$TMPFILE"

    if [ -d "$DEST" ]; then
        echo -e "${YELLOW}Removing existing ${DEST}${RESET}"
        rm -rf "$DEST"
    fi
    mv "/tmp/docx_redline_extract/docx_redline-${BRANCH}" "$DEST"
    rm -rf /tmp/docx_redline_extract
    return 0
}

if command -v git &>/dev/null; then
    echo "Downloading via git..."
    if git clone --depth 1 "https://github.com/${REPO}.git" "$DEST" 2>/dev/null; then
        :
    else
        echo ""
        echo -e "${YELLOW}Git clone failed (offline VPN, firewall, or GitHub blocked).${RESET}"
        echo -e "${YELLOW}Falling back to a ZIP download over HTTPS...${RESET}"
        echo ""
        if ! download_via_zip; then
            exit 1
        fi
    fi
else
    if ! download_via_zip; then
        exit 1
    fi
fi

echo -e "${GREEN}Downloaded.${RESET}"
echo ""

cd "$DEST"
bash setup.sh
