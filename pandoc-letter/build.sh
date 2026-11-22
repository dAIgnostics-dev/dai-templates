#!/usr/bin/env bash
# Build a dAIgnostics business letter from Markdown.
#   ./build.sh dopis.md              -> dopis.pdf
#   ./build.sh dopis.md dopis.docx   -> Word
set -euo pipefail

IN="${1:?usage: ./build.sh input.md [output.pdf|output.docx]}"
OUT="${2:-${IN%.md}.pdf}"
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
IN_DIR="$(cd "$(dirname "$IN")" && pwd)"

# The logo and the EU emblem are shared with the document templates (one
# copy of the registered logo in the repository): daignostics-logo.pdf and
# eu-logo.png from ../pandoc-document. TEXINPUTS lets XeLaTeX find them from
# any working directory; the resource path does the same for pandoc (DOCX).
SHARED="$HERE/../pandoc-document"
export TEXINPUTS="$HERE:$SHARED:${TEXINPUTS:-}"

COMMON=(--metadata-file="$HERE/company.yaml"
        --lua-filter="$HERE/letter.lua"
        --resource-path=".:$IN_DIR:$HERE:$SHARED")

# YAML front matter value of a top-level key (first match, quotes stripped)
front() {
  sed -n '1{/^---$/!q};1,/^\(---\|\.\.\.\)$/p' "$IN" | sed -n "2,\$s/^$1:[[:space:]]*//p" |
    head -n1 | tr -d "\"'" | tr '[:upper:]' '[:lower:]'
}

case "$OUT" in
  *.docx)
    LANG_REF=hr
    [[ "$(front lang)" == en* ]] && LANG_REF=en
    VARIANT="$(front letterhead)"
    VARIANT="${VARIANT:0:1}"
    [[ "$VARIANT" =~ ^[abc]$ ]] || VARIANT=c   # default: variant C (as letter.lua)
    pandoc "$IN" "${COMMON[@]}" \
      --reference-doc="$HERE/reference-letter-$LANG_REF-$VARIANT.docx" \
      -o "$OUT" ;;
  *)
    pandoc "$IN" "${COMMON[@]}" \
      --template="$HERE/daignostics-letter.latex" \
      --pdf-engine=xelatex \
      -o "$OUT" ;;
esac
echo "Wrote $OUT"
