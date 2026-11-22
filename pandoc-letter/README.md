# dAIgnostics business letter (pandoc → PDF / DOCX)

Letterhead template for outgoing business letters (*dopis*), built from the
same Markdown into PDF (XeLaTeX) or Word.

Fonts are the same as Sandi's document templates: TeX Gyre Heros (text) and
TeX Gyre Adventor (headings, subject) in PDF; Arial and Century Gothic in
Word. Colours and logo follow Marketing's brand rules (2026-11-22): text colours **#303030** (dark grey) and **#E21E3A** (red),
the official vector logo, and the slogan *Detect earlier. **Act smarter.**
For a life worth living.* with "Act smarter." in red. Contact labels
(T: / E: / W:) bold, data regular. The signatory's function is bold red.

Based on D. Frank's memorandum proposal (22.11.2026.). **Letterhead variant C**
(logo | red vertical line | contact, matching the e-mail signature) is the
default; variants A and B are kept and can be chosen per letter with
`letterhead:`. All three are implemented for both PDF and DOCX.

| Variant | Letterhead |
|---|---|
| `a` | logo + slogan left, contact right, red horizontal rule |
| `b` | centred, minimal: logo, slogan, short red rule, contact in one line |
| `c` (default) | logo + slogan, red vertical line, contact |

## Files

| File | Purpose |
|---|---|
| `daignostics-letter.latex` | PDF template (letterhead, address window, reference block, legal footer) |
| `letter.lua` | Shared filter: HR/EN labels; for DOCX builds the address/reference block, subject, closing, signature, enclosures |
| `company.yaml` | Company data for the letterhead and the statutory footer — **the only place to change them** |
| `make-reference-letter.py` | Regenerates `reference-letter-{hr,en}-{a,b,c}.docx` from `company.yaml` |
| `reference-letter-*.docx` | Pandoc reference documents per language and letterhead variant (styles, header, footer) |
| `build.sh` | Build helper: output extension decides PDF or DOCX; picks the reference document by `lang` and `letterhead` |
| `example-hr.md`, `example-en.md` (+ `.pdf`, `.docx`) | Worked examples and their outputs |

The logo and the EU emblem are **not copied** into this folder: the PDF uses
`../pandoc-document/daignostics-logo.pdf` and `eu-logo.png` (`build.sh` puts
that folder on `TEXINPUTS` and on the resource path), the reference documents
embed `../pandoc-docx/daignostics-logo.png`. These are the official logo files
(registered mark HR Z20251009): use only them, never redraw or recolour.

## Usage

```bash
./build.sh dopis.md               # → dopis.pdf
./build.sh dopis.md dopis.docx    # → Word (reference doc chosen by `lang`)
python3 make-reference-letter.py  # after editing company.yaml (needs python-docx, PyYAML)
```

## Metadata

```yaml
---
lang: hr                       # hr | en  (labels, hyphenation, footer)
letterhead: c                  # a | b | c  (default c)
recipient:                     # address window, max. 6 lines
  - "Primjer d.o.o."
  - "n/r gđa Ana Primjer"
  - "Ulica primjera 1"
  - "10000 Zagreb"
our-ref: "LET-26-001-R1-P-261122-Primjer d.o.o.-Naslov"   # code from Document coding.xlsx
your-ref: "Vaš dopis od 15. 11. 2026."                     # optional
place: "Zagreb"                # optional, printed before the date
date: "22. studenoga 2026."
contact: "Domagoj Frank"       # optional, + contact-email / contact-phone
subject: "Predmet dopisa"
opening: "Poštovana gospođo Primjer,"
closing: "S poštovanjem,"      # optional, default per language
signatory: "Domagoj Frank"
signatory-title: "direktor"
signature: "potpis.png"        # optional scanned signature
enclosures: ["Prilog 1 – …"]   # optional
cc: ["…"]                      # optional
eu-funding: true               # optional: EU emblem + funding sentence
funding: "…"                   # optional, overrides the default sentence
---
Body text in Markdown.
```

## Layout (best practice)

- A4, DIN 5008 form B: recipient in the window of a DL / C5/6 envelope
  (left window, 45–90 mm from the top), return-address line above it,
  reference block on the right (our ref., your ref., date, contact).
- Fold marks at 105 and 210 mm, punch mark at 148.5 mm (PDF only).
- Subject in the heading font, label "Predmet:" / "Subject:" (Croatian
  office practice).
- Page 1: full letterhead. Pages 2+: wordmark, reference and "Page X / Y".
- Footer on every page: statutory details required for business paper by
  the Croatian Companies Act (ZTD art. 21(4)–(5)): company name and seat,
  registry court and MBS, OIB, share capital, management, bank, IBAN.
- Closing and signature block are kept together; enclosure and cc lists
  are not split across pages.

## Differences between PDF and DOCX

- DOCX has no fold/punch marks.
- DOCX shows the EU emblem and funding sentence at the end of the letter,
  not in the footer (the footer lives in the reference doc and is the same
  for every letter).
- DOCX running header has no reference code, only "Stranica X / Y".
- Address/reference positions in DOCX are close to, but not exactly on, the
  DIN 5008 grid (about 1.5 mm higher than the PDF, the same for all three
  variants); check with a window envelope before mass mailing.
