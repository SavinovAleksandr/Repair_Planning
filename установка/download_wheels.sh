#!/bin/bash
# Download Windows-compatible wheels into установка/wheels/
set -euo pipefail
DIR="$(cd "$(dirname "$0")" && pwd)"
mkdir -p "$DIR/wheels"
python3 -m pip download openpyxl -d "$DIR/wheels" --only-binary=:all:
echo "Saved to $DIR/wheels"
ls -la "$DIR/wheels"
