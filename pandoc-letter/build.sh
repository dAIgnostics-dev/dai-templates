#!/usr/bin/env bash
# Build a dAIgnostics letter (memorandum) from one Markdown file.
#   ./build.sh letter.md              # → letter.pdf and letter.docx
#   ./build.sh letter.md pdf [out.pdf]
#   ./build.sh letter.md docx [out.docx]
set -euo pipefail

IN="${1:?usage: ./build.sh letter.md [pdf|docx|both] [output]}"
FMT="${2:-both}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
IN_DIR="$(cd "$(dirname "$IN")" && pwd)"
BASE="${IN%.md}"

# The logo is shared with the other templates (one source file):
# daignostics-logo.pdf from ../pandoc-document (via TEXINPUTS) for the PDF,
# ../pandoc-docx/daignostics-logo.png is embedded in reference.docx.
export TEXINPUTS="$HERE:$HERE/../pandoc-document:${TEXINPUTS:-}"

build_pdf() {
  pandoc "$IN" \
    --template="$HERE/daignostics-letter.latex" \
    --lua-filter="$HERE/letter.lua" \
    --pdf-engine=xelatex \
    --resource-path=".:$IN_DIR:$HERE" \
    -o "$1"
  echo "Wrote $1"
}

build_docx() {
  pandoc "$IN" \
    --reference-doc="$HERE/reference.docx" \
    --lua-filter="$HERE/letter.lua" \
    --resource-path=".:$IN_DIR:$HERE" \
    -o "$1"
  echo "Wrote $1"
}

case "$FMT" in
  pdf)  build_pdf  "${3:-$BASE.pdf}" ;;
  docx) build_docx "${3:-$BASE.docx}" ;;
  both) build_pdf "$BASE.pdf"; build_docx "$BASE.docx" ;;
  *) echo "unknown format: $FMT (pdf, docx or both)" >&2; exit 2 ;;
esac
