#!/usr/bin/env bash
# Build a DOCX from a markdown file using the dAIgnostics template.
#   ./build.sh input.md [output.docx]
set -euo pipefail

IN="${1:?usage: ./build.sh input.md [output.docx]}"
OUT="${2:-${IN%.md}.docx}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
IN_DIR="$(cd "$(dirname "$IN")" && pwd)"

# pandoc >= 3.8 renamed --highlight-style to --syntax-highlighting
if pandoc --help | grep -q -- '--syntax-highlighting'; then
  HL=(--syntax-highlighting="$HERE/daignostics.theme")
else
  HL=(--highlight-style="$HERE/daignostics.theme")
fi

pandoc "$IN" \
  --reference-doc="$HERE/reference.docx" \
  --lua-filter="$HERE/admonitions-docx.lua" \
  --lua-filter="$HERE/cover-docx.lua" \
  "${HL[@]}" \
  --resource-path=".:$IN_DIR:$HERE" \
  -o "$OUT"

echo "Wrote $OUT"
