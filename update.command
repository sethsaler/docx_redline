#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$DIR"

bash "$DIR/update.sh"
echo ""
echo "Press Enter to close."
read
