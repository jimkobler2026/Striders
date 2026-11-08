#!/bin/zsh
# Build the catalog workbook, install Mac Preview/Finder helpers, and save a .xlsm shell.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
XLSX="$ROOT/catalog/build/Strider_Media_Catalog.xlsx"
XLSM="$ROOT/catalog/Strider_Media_Catalog.xlsm"

python3 "$ROOT/scripts/build_catalog.py"
"$ROOT/scripts/install_mac_helpers.sh"

rm -f "$XLSM"
osascript <<EOF
set xlsxPath to POSIX file "$XLSX"
set xlsmPath to "$XLSM"

tell application "Microsoft Excel"
    activate
    open xlsxPath
    delay 1
    set wb to active workbook
    save workbook as wb filename xlsmPath file format macro enabled XML file format
    close wb saving no
end tell
EOF

echo
echo "Workbook ready (always open this file):"
echo "  $XLSM"
echo
echo "Do not open catalog/build/*.xlsx for daily use — it cannot store macros."
echo
echo "One-time VBA import in Excel:"
echo "  1. Open the .xlsm"
echo "  2. Tools → Macro → Visual Basic…"
echo "  3. File → Import File… → catalog/macros/CatalogMacros.bas"
echo "  4. Save (should stay .xlsm), enable macros when prompted"
echo "  5. Optional: Developer → Button, assign AppendCatalogRow / OpenSelectedInPreview / RevealSelectedInFinder"
