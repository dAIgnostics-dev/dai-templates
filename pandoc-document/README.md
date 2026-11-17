# dAIgnostics documentation template (pandoc → PDF)

A pandoc LaTeX template for software documentation, styled after
**daignostics.info**: clean white canvas, a single red accent (`#E63946`),
geometric Century-Gothic-style headings (TeX Gyre Adventor), and the
**Co-funded by the European Union** emblem on the cover and in the footer of
every page.

## Files

| File                | Purpose                                                        |
|---------------------|----------------------------------------------------------------|
| `daignostics.latex` | The pandoc template (use with `--template`)                    |
| `admonitions.lua`   | Filter turning `::: note` fenced divs into styled callout boxes |
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

Keep `eu-logo.png` next to your document, or point the `eu-logo:` metadata key
at another path. If the logo file is missing the build still succeeds — it just
omits the emblem.

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
# header-left: "Your Project"             # overrides the dAIgnostics wordmark
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
  `daignostics.latex` — edit `\definecolor{dgred}{HTML}{E63946}` to retheme.
- **Body font** defaults to TeX Gyre Heros. Override per-document with
  `mainfont:` / `monofont:` in YAML, or globally in the template.
- **Logo placement** is the footer (right) on content pages and the lower-right
  of the cover. Change the `\fancyfoot[R]` / cover `minipage` blocks to move it.
