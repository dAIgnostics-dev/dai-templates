# dAIgnostics documentation template (pandoc → PDF)

A pandoc LaTeX template for software documentation, styled after
**daignostics.info**: clean white canvas, the dAIgnostics logo on the cover
and in the running header, the logo colours (red `#E21E3A`, ink `#303030`),
geometric Century-Gothic-style headings (TeX Gyre Adventor), and the
**Co-funded by the European Union** emblem on the cover and in the footer of
every page.

## Files

| File                | Purpose                                                        |
|---------------------|----------------------------------------------------------------|
| `daignostics.latex` | The pandoc template (use with `--template`)                    |
| `admonitions.lua`   | Filter turning `::: note` fenced divs into styled callout boxes |
| `daignostics-logo.pdf` | Official dAIgnostics logo (vector), cover + running header — see *Logo* below |
| `daignostics-logo.svg` | Marketing's vector original of the logo (source of the `.pdf` / `.png`) |
| `eu-logo.png`       | The EU emblem, trimmed; referenced by every page               |
| `build.sh`          | One-line build helper                                          |
| `example.md` / `example.pdf` | Worked example and its output                         |

## Requirements

- `pandoc` (tested on 3.1)
- A TeX Live install with **XeLaTeX** and the `tex-gyre` fonts (Adventor,
  Heros), plus `titlesec`, `fancyhdr`, `tcolorbox`, `lastpage`, `enumitem`.
  On Ubuntu: `texlive-xetex texlive-fonts-extra texlive-latex-extra`.

## Usage

```bash
./build.sh mydoc.md            # → mydoc.pdf
./build.sh mydoc.md out.pdf
```

Or call pandoc directly:

```bash
pandoc mydoc.md \
  --template=daignostics.latex \
  --lua-filter=admonitions.lua \
  --pdf-engine=xelatex \
  --highlight-style=tango \
  -o mydoc.pdf
```

`build.sh` puts the template folder on `TEXINPUTS`, so `daignostics-logo.pdf`
and `eu-logo.png` are found whichever directory you build from. When calling
pandoc directly from elsewhere, do the same
(`TEXINPUTS=/path/to/pandoc-document: pandoc …`), keep the files next to your
document, or point the `eu-logo:` metadata key at another path. Missing files
never break the build: without `daignostics-logo.pdf` the cover and header show
the text wordmark, without `eu-logo.png` the emblem is omitted.

## Logo

`daignostics-logo.pdf` is the official logo (registered figurative trademark
HR Z20251009), rendered from Marketing's vector original, kept here as
`daignostics-logo.svg` (Marketing: *Visuals/daignostics-vector.svg*). The
Word template uses a PNG rendering of the same file
(`../pandoc-docx/daignostics-logo.png`). The logo is used exactly as supplied
— only scaled by height (8 mm on the cover, 4.5 mm in the header), never
redrawn, stretched or recoloured. To replace it, replace the files; do not
edit them.

## Metadata keys (YAML front matter)

```yaml
---
title:    "Neural Signal Analysis API"   # required for the cover
subtitle: "Integration & Operations Guide"
doctype:  "Software Documentation"        # eyebrow label above the title
author:   ["Name One", "Name Two"]
project:  "Sample Project"                # shown top-right of cover + footer-left
version:  "1.4.0"
date:     "June 2026"
toc: true
toc-depth: 2
numbersections: true
lang: en                                  # use `hr` for Croatian hyphenation
eu-logo: "eu-logo.png"                    # path to the EU emblem
# confidential: true                      # stamps "Confidential" in the footer
# header-left: "Your Project"             # overrides the logo in the running header
# funding: "Custom funding acknowledgement text for the cover."
---
```

All keys are optional except `title` (needed to draw the cover).

## Callout boxes

With `admonitions.lua` enabled, fenced divs become styled boxes:

```markdown
::: note
A neutral note (slate spine).
:::

::: tip
A tip (blue spine). `info` looks the same.
:::

::: {.warning title="Custom heading"}
A warning (red spine). `important` looks the same.
:::
```

Classes: `note`, `tip`, `info`, `warning`, `important`. The optional
`title="..."` attribute overrides the heading.

## Customizing

- **Accent colour, fonts, margins** live in the block marked
  `dAIgnostics software-documentation theme` near the top of
  `daignostics.latex`. `dgred` (`E21E3A`) and `dgink` (`303030`) are the logo
  colours, the same values the e-mail signature and letterhead rules use —
  keep them in sync with `../pandoc-docx/make-reference.py`.
- **Body font** defaults to TeX Gyre Heros. Override per-document with
  `mainfont:` / `monofont:` in YAML, or globally in the template.
- **EU emblem placement** is the footer (right) on content pages and the
  lower-right of the cover. Change the `\fancyfoot[R]` / cover `minipage`
  blocks to move it.

## Page breaks

The *Page-break hygiene* block (after the `\titlespacing` lines) keeps headings
and tables from being stranded at the bottom of a page:

- `\section`, `\subsection`, `\subsubsection` require 26 / 13 / 12 % of the
  text height to be free, otherwise they start on the next page.
- A table (`longtable`) needs 11 % free, except directly after a heading:
  there it stays with the heading (the heading already reserved the room).
- Widows, orphans and page breaks after a hyphen are forbidden.

If a page now ends with more white space than you like, the percentages are in
`\Needspace*{…\textheight}`.
