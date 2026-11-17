#!/usr/bin/env bash
# Build a PDF from a markdown file using the dAIgnostics template.
#   ./build.sh input.md [output.pdf]
set -euo pipefail

IN="${1:?usage: ./build.sh input.md [output.pdf]}"
OUT="${2:-${IN%.md}.pdf}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

pandoc "$IN" \
  --template="$HERE/daignostics.latex" \
  --lua-filter="$HERE/admonitions.lua" \
  --pdf-engine=xelatex \
  --highlight-style=tango \
  --resource-path=".:$HERE" \
  -o "$OUT"

echo "Wrote $OUT"
