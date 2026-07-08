#!/usr/bin/env bash
# Render a .pptx to per-slide JPGs for visual QA.
# Usage: render_qa.sh /absolute/path/to/deck.pptx
# Output: qa-1.jpg, qa-2.jpg, ... written next to the .pptx.
set -euo pipefail

PPTX="${1:?usage: render_qa.sh /path/to/deck.pptx}"
DIR="$(cd "$(dirname "$PPTX")" && pwd)"
BASE="$(basename "$PPTX" .pptx)"

# Find the pptx skill's bundled soffice wrapper if available (handles sandboxed
# LibreOffice profiles); otherwise fall back to plain `soffice`.
SOFFICE_WRAP="$(find "$HOME/.claude" -path '*skills/pptx/scripts/office/soffice.py' 2>/dev/null | head -1 || true)"

cd "$DIR"
if [[ -n "$SOFFICE_WRAP" ]]; then
  python3 "$SOFFICE_WRAP" --headless --convert-to pdf "$PPTX" >/dev/null 2>&1
else
  soffice --headless --convert-to pdf "$PPTX" >/dev/null 2>&1
fi

pdftoppm -jpeg -r 120 "$DIR/$BASE.pdf" "$DIR/qa" >/dev/null 2>&1
rm -f "$DIR/$BASE.pdf"
echo "Rendered:"
ls "$DIR"/qa-*.jpg
