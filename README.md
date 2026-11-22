# dai-templates

Shared templates for dAIgnostics documents and projects.

| Template | What it is |
|---|---|
| [`pandoc-document/`](pandoc-document/) | Pandoc → PDF template (XeLaTeX) for software documentation in the dAIgnostics style: official logo and logo colours, TeX Gyre fonts, callout boxes, and the "Co-funded by the European Union" emblem on the cover and footer. See its [README](pandoc-document/README.md) and [example PDF](pandoc-document/example.pdf). |
| [`pandoc-docx/`](pandoc-docx/) | Pandoc → DOCX (Word) version of the same template: builds the same Markdown into a `.docx` with a matching cover page, Century Gothic headings, callout boxes and the EU emblem in the footer. See its [README](pandoc-docx/README.md) and [example DOCX](pandoc-docx/example.docx). |
| [`pandoc-letter/`](pandoc-letter/) | Letter / memorandum on the dAIgnostics letterhead: one Markdown file builds both a PDF (XeLaTeX) and a DOCX with the logo, recipient block, subject, signature and the company details in the footer. See its [README](pandoc-letter/README.md), [example PDF](pandoc-letter/example.pdf) and [example DOCX](pandoc-letter/example.docx). |

## Quick start

```bash
git clone https://github.com/daignostics-dev/dai-templates.git
cd dai-templates/pandoc-document
./build.sh example.md   # → example.pdf

# Same document as Word:
cd ../pandoc-docx
./build.sh ../pandoc-document/example.md example.docx   # → example.docx

# Letter / memorandum, PDF + DOCX from one file:
cd ../pandoc-letter
./build.sh example.md   # → example.pdf, example.docx
```

Originally extracted from the internal `k-wiki` repository (`pandoc-template/`).
