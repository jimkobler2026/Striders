#!/bin/zsh
# Install AppleScript helpers so Excel VBA can open Preview / reveal in Finder / caption.
set -euo pipefail

ROOT="$(cd "$(dirname "$0")/.." && pwd)"
SRC="$ROOT/catalog/macros"
DEST="$HOME/Library/Application Scripts/com.microsoft.Excel"
ENV_FILE="$ROOT/../Virtual_Nexus/.env"
PY="$(command -v python3)"

mkdir -p "$DEST"

osacompile -o "$DEST/StriderOpenPreview.scpt" "$SRC/StriderOpenPreview.applescript"
osacompile -o "$DEST/StriderRevealFinder.scpt" "$SRC/StriderRevealFinder.applescript"
osacompile -o "$DEST/StriderFinderMeta.scpt" "$SRC/StriderFinderMeta.applescript"
osacompile -o "$DEST/StriderCaption.scpt" "$SRC/StriderCaption.applescript"

# Shell bridge used by StriderCaption.scpt (keeps project paths out of VBA).
cat > "$DEST/StriderCaption.sh" <<EOF
#!/bin/zsh
set -euo pipefail
ROOT="$ROOT"
PY="$PY"
ENV_FILE="$ENV_FILE"
PHOTO="\${1:-}"
PROVIDER="\${2:-google}"
if [[ -z "\$PHOTO" ]]; then
  echo "ERROR\tNo photo path" >&2
  exit 2
fi
ARGS=(scripts/caption_photo.py "\$PHOTO" --provider "\$PROVIDER" --machine)
if [[ -f "\$ENV_FILE" ]]; then
  ARGS+=(--env-file "\$ENV_FILE")
elif [[ -f "\$ROOT/.env" ]]; then
  ARGS+=(--env-file "\$ROOT/.env")
fi
cd "\$ROOT"
exec "\$PY" "\${ARGS[@]}"
EOF
chmod +x "$DEST/StriderCaption.sh"

echo "Installed:"
echo "  $DEST/StriderOpenPreview.scpt"
echo "  $DEST/StriderRevealFinder.scpt"
echo "  $DEST/StriderFinderMeta.scpt"
echo "  $DEST/StriderCaption.scpt"
echo "  $DEST/StriderCaption.sh"
echo
echo "Restart Microsoft Excel if it was open, then open the .xlsm catalog."
echo "Captions default to Gemini Flash (google). Optional: anthropic / openai."
