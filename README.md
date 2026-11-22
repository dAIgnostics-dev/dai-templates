# dai-templates

Shared templates for dAIgnostics documents and projects.

| Template | What it is |
|---|---|
| [`pandoc-document/`](pandoc-document/) | Pandoc → PDF template (XeLaTeX) for software documentation in the dAIgnostics style: official logo and logo colours, TeX Gyre fonts, callout boxes, and the "Co-funded by the European Union" emblem on the cover and footer. See its [README](pandoc-document/README.md) and [example PDF](pandoc-document/example.pdf). |
| [`pandoc-docx/`](pandoc-docx/) | Pandoc → DOCX (Word) version of the same template: builds the same Markdown into a `.docx` with a matching cover page, Century Gothic headings, callout boxes and the EU emblem in the footer. See its [README](pandoc-docx/README.md) and [example DOCX](pandoc-docx/example.docx). |
| [`pandoc-letter/`](pandoc-letter/) | Business letter (*dopis*) on the dAIgnostics letterhead, PDF and DOCX from one Markdown file: logo, slogan and contact (letterhead variant C by default, A/B selectable), DIN 5008 address window and reference block, signature, and the company's statutory details (ZTD art. 21) from `company.yaml` in the footer. See its [README](pandoc-letter/README.md) and the examples ([PDF](pandoc-letter/example-hr.pdf), [DOCX](pandoc-letter/example-hr.docx)). |

## Quick start

```bash
git clone https://github.com/daignostics-dev/dai-templates.git
cd dai-templates/pandoc-document
./build.sh example.md   # → example.pdf

# Same document as Word:
cd ../pandoc-docx
./build.sh ../pandoc-document/example.md example.docx   # → example.docx

# Business letter (memorandum), PDF or DOCX from the same file:
cd ../pandoc-letter
./build.sh example-hr.md                    # → example-hr.pdf
./build.sh example-hr.md example-hr.docx    # → Word
```

Originally extracted from the internal `k-wiki` repository (`pandoc-template/`).
