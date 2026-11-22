#!/usr/bin/env python3
"""Generate reference.docx for the dAIgnostics letter (memorandum) template.

    python3 make-reference.py              # writes reference.docx

Reuses ../pandoc-docx/make-reference.py (palette, fonts, theme, body and
table styles) and adds what a letter needs: the styles used by letter.lua
(recipient, date, subject, closing, signature, enclosures), a letterhead on
page 1 (large logo + red rule), a small logo with "page / total" on the
following pages, and the company details in the footer of every page.

Requires: pandoc on PATH, python-docx (tested with 1.2.0).
"""
import argparse
import importlib.util
import os
import subprocess
import sys
import tempfile

from docx import Document
from docx.shared import Cm, Twips

HERE = os.path.dirname(os.path.abspath(__file__))
sys.dont_write_bytecode = True   # no __pycache__ in ../pandoc-docx
_spec = importlib.util.spec_from_file_location(
    "dg_reference", os.path.join(HERE, "..", "pandoc-docx", "make-reference.py"))
B = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(B)

# ---- Company details (same as the footer in daignostics-letter.latex) ----
COMPANY = "DAIGNOSTICS d.o.o."
LINE1 = ["Bobovje 52G, 49000 Bobovje (grad Krapina), Hrvatska", "www.daignostics.info"]
LINE2 = ["OIB: 58352432266", "IBAN: HR4324840081502035263", "contact@daignostics.info"]

# ---- Page geometry (A4, as daignostics-letter.latex) -----------------------
MARGIN_T, MARGIN_B = 1701, 1814        # 3.0 cm / 3.2 cm
HEADER_D, FOOTER_D = 680, 454          # 1.2 cm / 0.8 cm


def letter_styles():
    ps, sz = B.pstyle, B.sz
    left = '<w:jc w:val="left"/>'
    return [
        ps("LetterRecipient", "Letter Recipient",
           '<w:spacing w:before="170" w:after="0" w:line="264" w:lineRule="auto"/>' + left, ""),
        ps("LetterDate", "Letter Date",
           '<w:spacing w:before="567" w:after="0"/><w:jc w:val="right"/>', ""),
        ps("LetterSubject", "Letter Subject",
           '<w:keepNext/><w:spacing w:before="397" w:after="340"/>' + left,
           B.head_fonts() + sz(10.5)),
        ps("LetterOpening", "Letter Opening",
           '<w:keepNext/><w:spacing w:before="0" w:after="140"/>' + left, ""),
        ps("LetterClosing", "Letter Closing",
           '<w:keepNext/><w:keepLines/><w:spacing w:before="280" w:after="255"/>' + left, ""),
        ps("LetterSigner", "Letter Signer",
           '<w:keepNext/><w:spacing w:before="0" w:after="0"/>' + left, ""),
        ps("LetterSignerTitle", "Letter Signer Title",
           '<w:keepNext/><w:spacing w:before="0" w:after="0"/>' + left,
           f'<w:color w:val="{B.MUTED}"/>'),
        ps("LetterSignature", "Letter Signature",
           '<w:tabs><w:tab w:val="left" w:leader="underscore" w:pos="4400"/></w:tabs>'
           '<w:spacing w:before="620" w:after="0"/>' + left, ""),
        ps("LetterEnclosuresTitle", "Letter Enclosures Title",
           '<w:keepNext/><w:spacing w:before="454" w:after="40"/>' + left, sz(9)),
        ps("LetterEnclosure", "Letter Enclosure",
           '<w:spacing w:before="0" w:after="20"/>' + left, sz(9)),
        B.cstyle("LetterLabel", "Letter Label", f'<w:color w:val="{B.MUTED}"/>'),
        ps("LetterFooter", "Letter Footer",
           f'<w:pBdr><w:top w:val="single" w:sz="4" w:space="6" w:color="{B.LINE}"/></w:pBdr>'
           '<w:spacing w:before="0" w:after="0" w:line="230" w:lineRule="auto"/><w:jc w:val="center"/>',
           f'<w:color w:val="{B.MUTED}"/>' + sz(7.5), based="Footer"),
    ]


def logo_or_wordmark(doc, p, logo, height_cm):
    """Official logo (registered mark HR Z20251009), scaled by height only."""
    if os.path.exists(logo):
        p.add_run().add_picture(logo, height=Cm(height_cm))
    else:
        wm = '<w:rStyle w:val="Wordmark"/>'
        for xml in (B.text_run("d", wm), B.text_run("AI", wm + f'<w:color w:val="{B.RED}"/>'),
                    B.text_run("gnostics", wm)):
            B.append_xml(p._p, xml)


def company_footer(doc, footer):
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0]
    p.style = doc.styles["Letter Footer"]
    sep = B.text_run(" | ", f'<w:color w:val="{B.RED}"/>')
    runs = [B.text_run(COMPANY, f'<w:b/><w:color w:val="{B.INK}"/>')]
    for part in LINE1:
        runs += [sep, B.text_run(part)]
    runs.append(f'<w:r {B.NSDECL}><w:br/></w:r>')
    for i, part in enumerate(LINE2):
        runs += ([sep] if i else []) + [B.text_run(part)]
    for xml in runs:
        B.append_xml(p._p, xml)


def build_header_footer(doc, logo):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Twips(B.PAGE_W), Twips(B.PAGE_H)
    sec.left_margin = sec.right_margin = Twips(B.MARGIN_LR)
    sec.top_margin, sec.bottom_margin = Twips(MARGIN_T), Twips(MARGIN_B)
    sec.header_distance, sec.footer_distance = Twips(HEADER_D), Twips(FOOTER_D)
    for old in sec._sectPr.findall(B.qn("w:pgNumType")):
        sec._sectPr.remove(old)
    sec.different_first_page_header_footer = True

    # page 1: letterhead - large logo over the red rule (Header style)
    fh = sec.first_page_header
    fh.is_linked_to_previous = False
    p = fh.paragraphs[0]
    p.style = doc.styles["Header"]
    logo_or_wordmark(doc, p, logo, 0.9)

    # following pages: small logo, page / total on the right
    hdr = sec.header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0]
    p.style = doc.styles["Header"]
    logo_or_wordmark(doc, p, logo, 0.45)
    B.append_xml(p._p, f'<w:r {B.NSDECL}><w:tab/></w:r>')
    muted = f'<w:color w:val="{B.MUTED}"/>'
    for xml in (B.field_runs("PAGE", muted, "2") + [B.text_run(" / ", muted)]
                + B.field_runs("NUMPAGES", muted, "2")):
        B.append_xml(p._p, xml)

    company_footer(doc, sec.first_page_footer)
    company_footer(doc, sec.footer)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("-o", "--output", default=os.path.join(HERE, "reference.docx"))
    ap.add_argument("--brand-logo",
                    default=os.path.join(HERE, "..", "pandoc-docx", "daignostics-logo.png"),
                    help="dAIgnostics logo (default: ../pandoc-docx/daignostics-logo.png; "
                         "text wordmark if the file is missing)")
    a = ap.parse_args()

    with tempfile.TemporaryDirectory() as td:
        base = os.path.join(td, "base.docx")
        subprocess.run(["pandoc", "-o", base, "--print-default-data-file", "reference.docx"],
                       check=True)
        doc = Document(base)
    B.set_doc_defaults(doc)
    B.replace_styles(doc, B.build_styles() + [B.table_style()] + letter_styles())
    B.patch_part_blob(doc, "/word/theme/theme1.xml", B.patch_theme)
    B.patch_part_blob(doc, "/word/fontTable.xml", B.patch_font_table)
    build_header_footer(doc, a.brand_logo)
    cp = doc.core_properties
    cp.title, cp.author, cp.comments = "dAIgnostics letter reference document", "dAIgnostics", \
        "Generated by make-reference.py - pandoc --reference-doc"
    doc.save(a.output)
    print(f"Wrote {a.output}")


if __name__ == "__main__":
    sys.exit(main())
