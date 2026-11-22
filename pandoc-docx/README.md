# dAIgnostics documentation template (pandoc → DOCX)

The Word counterpart of [`../pandoc-document/`](../pandoc-document/) (pandoc →
PDF). The **same Markdown** (same YAML front matter, same `::: note` callouts)
builds to a `.docx` styled like the PDF: the dAIgnostics logo on the cover
and in the running header, the logo colours (red `#E21E3A`, ink `#303030`),
geometric Century Gothic headings, a footer with the **Co-funded by the
European Union** emblem, and shaded callout boxes.

## Files

| File                    | Purpose                                                                      |
|-------------------------|------------------------------------------------------------------------------|
| `reference.docx`        | Pandoc reference document: styles, page setup, header/footer, logo, EU emblem |
| `make-reference.py`     | Regenerates `reference.docx` and `daignostics.theme` (python-docx + pandoc)  |
| `cover-docx.lua`        | Builds the cover page from metadata; handles `toc`, `numbersections`, `abstract`, `confidential`, `lang: hr` |
| `admonitions-docx.lua`  | Turns `::: note/tip/info/warning/important` divs into callout boxes         |
| `daignostics.theme`     | Code highlighting theme (tango, background matched to the code box)         |
| `daignostics-logo.png`  | Official dAIgnostics logo (raster), cover + running header — see the PDF template's *Logo* section |
| `eu-logo.png`           | The EU emblem (copy of `../pandoc-document/eu-logo.png`)                     |
| `build.sh`              | One-line build helper                                                        |
| `example.docx`          | Output for `../pandoc-document/example.md`                                    |

## Requirements

- `pandoc` 3.x (tested on 3.9; `build.sh` also handles pre-3.8 flag names).
- To regenerate `reference.docx` only: `python3` with `python-docx` (tested 1.2.0).

## Usage

```bash
./build.sh mydoc.md              # → mydoc.docx
./build.sh mydoc.md out.docx
./build.sh ../pandoc-document/example.md example.docx
```

Or call pandoc directly (filter order matters: callouts first):

```bash
pandoc mydoc.md \
  --reference-doc=reference.docx \
  --lua-filter=admonitions-docx.lua \
  --lua-filter=cover-docx.lua \
  --syntax-highlighting=daignostics.theme \
  -o mydoc.docx
```

Turn the TOC on with `toc: true` in the YAML, **not** with `--toc` on the
command line: pandoc places a command-line TOC before the cover.

When opening the file, Word may ask to *update fields that may refer to other
files* — answer **Yes**. That fills the table of contents with page numbers and
refreshes the title/project in the header and footer. (LibreOffice shows the
TOC entries without page numbers until you right-click → *Update Index*.)

## Metadata keys (YAML front matter)

Same keys as the PDF template:

```yaml
---
title:    "Neural Signal Analysis API"   # required for the cover
subtitle: "Integration & Operations Guide"
doctype:  "Software Documentation"        # eyebrow label above the title
author:   ["Name One", "Name Two"]
project:  "Sample Project"                # cover top-right + footer-left
version:  "1.4.0"
date:     "June 2026"
toc: true
toc-depth: 2
numbersections: true
lang: en                                  # `hr`: Croatian labels, TOC title, spell-check language
# confidential: true                      # "CONFIDENTIAL" in the footer instead of the project
# funding: "Custom funding acknowledgement text for the cover."   (max. 255 characters)
# toc-title: "Contents"
---
```

`eu-logo:` and `header-left:` are accepted but ignored: in Word the emblem and
the header logo are part of `reference.docx` (see *Customizing*). The cover
logo is `daignostics-logo.png` next to `cover-docx.lua`; without that file the
cover and (after rerunning `make-reference.py`) the header show the text
wordmark.

## Callout boxes

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

Each callout becomes a title paragraph plus body paragraphs in the styles
*Callout Note / Tip / Warning (Title)*; they share one border and shading
definition, so Word draws them as a single box. Lists inside a callout are
rendered as paragraphs with a `•` / `1.` marker so the box stays intact.

## Fonts

Word machines rarely have the TeX Gyre fonts the PDF uses, so the reference
document uses the Office fonts they imitate:

| Role      | PDF (XeLaTeX)          | DOCX            | Fallback if missing                         |
|-----------|------------------------|-----------------|---------------------------------------------|
| Headings, cover, wordmark | TeX Gyre Adventor | **Century Gothic** (theme *major* font) | another sans (font table: swiss, alt. name URW Gothic) |
| Body      | TeX Gyre Heros         | **Arial** (theme *minor* font) | Liberation Sans / Arimo (metric-compatible) |
| Code      | DejaVu Sans Mono       | **Consolas**    | a monospace font (font table: modern, fixed pitch) |

Century Gothic ships with Windows and with Microsoft Office on macOS. Styles
name the font explicitly *and* point at the theme fonts, so switching the
document theme fonts in Word restyles everything consistently. On Linux /
LibreOffice without Century Gothic, map it to URW Gothic (same Avant-Garde
design) with a fontconfig alias, e.g. in `~/.config/fontconfig/fonts.conf`:

```xml
<match target="pattern"><test name="family"><string>Century Gothic</string></test>
  <edit name="family" mode="assign" binding="strong"><string>URW Gothic</string></edit></match>
```

## Page breaks

Heading 1–6 have *Keep with next* and *Keep lines together*, so a heading
never ends a page alone; Table Caption has *Keep with next*, so a caption
stays with its table. (The PDF template does the same with `\Needspace`.)

## Differences from the PDF template

- **Page numbers** count the cover (the cover itself shows none), so the first
  content page is `2 / N`. The PDF starts at `1` after the cover; Word's
  `SECTIONPAGES` / `{= NUMPAGES - 1}` could mimic that but LibreOffice and
  other viewers do not evaluate them.
- **Header/footer text** (title, project, "Confidential") are fields fed by
  document properties; Word refreshes them on open. "Confidential" is muted
  grey rather than red small caps.
- **TOC page numbers** appear only after Word updates the field (see above).
- **Tables** keep pandoc's equal column widths instead of LaTeX's natural
  widths; Word centres them, LibreOffice aligns them left.
- **Callouts** cannot contain a code block without breaking the box into two
  (the code block has its own red-spined box), and nested lists are flattened.
- **Cover without `title`**: no cover is generated, but page 1 still uses the
  first-page footer (funding text + large emblem) and has no running header.
- `header-left` and `eu-logo` are not per-document (they live in `reference.docx`).

## Customizing

Edit the constants at the top of `make-reference.py` (palette, fonts, margins)
and rerun it:

```bash
python3 make-reference.py                               # → reference.docx, daignostics.theme
python3 make-reference.py --page-label "Page {page} of {total}"
python3 make-reference.py --logo other-emblem.png      # EU emblem in the footer
python3 make-reference.py --brand-logo logo.png         # header logo (default daignostics-logo.png)
```

You can also open `reference.docx` in Word, change styles (Heading 1, Body
Text, Callout Note, Cover Title, …) and save it — but then the changes are
lost the next time someone runs the generator.
