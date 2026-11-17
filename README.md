# dai-templates

Shared templates for dAIgnostics documents and projects.

| Template | What it is |
|---|---|
| [`pandoc-document/`](pandoc-document/) | Pandoc → PDF template (XeLaTeX) for software documentation in the dAIgnostics style: red accent, TeX Gyre fonts, callout boxes, and the "Co-funded by the European Union" emblem on the cover and footer. See its [README](pandoc-document/README.md) and [example PDF](pandoc-document/example.pdf). |

## Quick start

```bash
git clone https://github.com/daignostics-dev/dai-templates.git
cd dai-templates/pandoc-document
./build.sh example.md   # → example.pdf
```

Originally extracted from the internal `k-wiki` repository (`pandoc-template/`).
