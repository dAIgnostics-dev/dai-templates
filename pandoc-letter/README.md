# dAIgnostics letter / memorandum (pandoc → PDF + DOCX)

One Markdown file → a letter on the dAIgnostics letterhead, as **PDF**
(XeLaTeX) and as **Word** (DOCX), with the same layout:

- page 1: official logo over a red rule (letterhead); later pages: small logo
  and `page / total`
- recipient block (first line bold) and *N/R*, place and date on the right,
  *Predmet* (subject)
- opening, body, closing, *DAIGNOSTICS d.o.o.*, signer, signer title and a
  signature line (kept together on one page), optional enclosures
- footer on every page: company name, address, web, OIB, IBAN, e-mail (as in
  the official memorandum)

Fonts and colours are those of the other templates: TeX Gyre Heros + Adventor
in the PDF, Arial + Century Gothic in Word; logo red `#E21E3A`, ink `#303030`.

## Files

| File                        | Purpose                                                              |
|-----------------------------|----------------------------------------------------------------------|
| `daignostics-letter.latex`  | Pandoc LaTeX template (PDF)                                           |
| `letter.lua`                | Filter for both outputs: labels by language, defaults; builds the letter blocks for DOCX |
| `reference.docx`            | Pandoc reference document (styles, letterhead, footer)               |
| `make-reference.py`         | Regenerates `reference.docx` (reuses `../pandoc-docx/make-reference.py`) |
| `build.sh`                  | Builds PDF and/or DOCX                                               |
| `example.md` / `.pdf` / `.docx` | Example letter and its outputs                                    |

The logo is not copied here: the PDF uses `../pandoc-document/daignostics-logo.pdf`
(`build.sh` puts that folder on `TEXINPUTS`), `reference.docx` embeds
`../pandoc-docx/daignostics-logo.png`. Keep the repository layout intact, or
copy those files next to this folder.

## Usage

```bash
./build.sh letter.md              # → letter.pdf and letter.docx
./build.sh letter.md pdf          # PDF only
./build.sh letter.md docx out.docx
```

Requirements as for `../pandoc-document` (pandoc 3.x, XeLaTeX, TeX Gyre fonts)
and, only to regenerate `reference.docx`, python-docx (tested 1.2.0).

## Metadata keys (YAML front matter)

```yaml
---
to:                                   # recipient, one line per item; first line bold
  - "Naziv primatelja d.o.o."
  - "Ulica i kućni broj 1"
  - "10000 Zagreb"
attention: "Ime Prezime"              # N/R (Attn.)
place: "Bobovje"                      # default: Bobovje
date: "22.11.2026."                   # default: today (22.11.2026. / 22 November 2026)
subject: "Predmet dopisa"
opening: "Poštovani,"                 # default by language
closing: "S poštovanjem,"             # default by language
organisation: "DAIGNOSTICS d.o.o."    # default
signer: "Ime Prezime"                 # default: first `author`
signer-title: "Funkcija"
signature-line: true                  # "Potpis: ____"; false to omit
enclosures: ["Prilog 1", "Prilog 2"]  # "Prilozi:" list at the end
lang: hr                              # default hr; `en` for English labels and date
---
```

All keys are optional. The body is ordinary Markdown (paragraphs, emphasis,
links, lists, tables). After *Poštovani,* the body starts with a lower-case
letter, as usual in Croatian letters.

## Company details

The footer text is in two places — keep them identical:

- `daignostics-letter.latex`: `\dgcompany`
- `make-reference.py`: `COMPANY`, `LINE1`, `LINE2` (then rerun it)
