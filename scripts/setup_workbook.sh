#!/bin/zsh
# Build a fresh .xlsm shell from catalog/build/*.xlsx and install Mac helpers.
# WARNING: Overwrites catalog/Strider_Media_Catalog.xlsm (macros/data lost).
# Prefer the hand-tuned blank .xlsm for day-to-day; use this only for a cold rebuild.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
XLSX="$ROOT/catalog/build/Strider_Media_Catalog.xlsx"
XLSM="$ROOT/catalog/Strider_Media_Catalog.xlsm"
FORCE=0

for arg in "$@"; do
  case "$arg" in
    --force|-f) FORCE=1 ;;
    --help|-h)
      cat <<'USAGE'
Usage: ./scripts/setup_workbook.sh [--force]

Rebuilds catalog/build/*.xlsx, installs Mac helpers, then writes a new
catalog/Strider_Media_Catalog.xlsm shell (no VBA modules yet).

If the live .xlsm already exists, this script refuses unless --force is passed.
Day-to-day: open the existing .xlsm. Helpers only: ./scripts/install_mac_helpers.sh
USAGE
      exit 0
      ;;
    *)
      echo "Unknown option: $arg (try --help)" >&2
      exit 2
      ;;
  esac
done

if [[ -e "$XLSM" && "$FORCE" -ne 1 ]]; then
  echo "Refusing to overwrite existing workbook:" >&2
  echo "  $XLSM" >&2
  echo >&2
  echo "That file holds your macros and catalog data." >&2
  echo "For helpers only:  ./scripts/install_mac_helpers.sh" >&2
  echo "For a blank copy:  catalog/build/Strider_Media_Catalog_blank.xlsm" >&2
  echo "To rebuild anyway: ./scripts/setup_workbook.sh --force" >&2
  exit 1
fi

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
echo "  4. Paste Entry / VideoEntry SelectionChange snippets from catalog/macros/*.txt"
echo "  5. Save (should stay .xlsm), enable macros when prompted"
