#!/bin/zsh
# Install AppleScript helpers so Excel VBA can open Preview / reveal in Finder.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
SRC="$ROOT/catalog/macros"
DEST="$HOME/Library/Application Scripts/com.microsoft.Excel"

mkdir -p "$DEST"

osacompile -o "$DEST/StriderOpenPreview.scpt" "$SRC/StriderOpenPreview.applescript"
osacompile -o "$DEST/StriderRevealFinder.scpt" "$SRC/StriderRevealFinder.applescript"
osacompile -o "$DEST/StriderFinderMeta.scpt" "$SRC/StriderFinderMeta.applescript"

echo "Installed:"
echo "  $DEST/StriderOpenPreview.scpt"
echo "  $DEST/StriderRevealFinder.scpt"
echo "  $DEST/StriderFinderMeta.scpt"
echo
echo "Restart Microsoft Excel if it was open, then open the .xlsm catalog."
