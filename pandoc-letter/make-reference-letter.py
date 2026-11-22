#!/usr/bin/env python3
"""Generate reference-letter-{hr,en}-{a,b,c}.docx.

    python3 make-reference-letter.py

Starts from pandoc's default reference.docx and restyles it for the
dAIgnostics business letter: A4, same margins as daignostics-letter.latex,
first-page letterhead in variant A, B or C (as the `letterhead:` key of the
PDF template), running header on pages 2+ (logo, "Stranica X / Y"), and the
statutory company details (ZTD art. 21) in the footer of every page. Company
data are read from company.yaml, so rerun this whenever that file changes.

Requires: pandoc on PATH, python-docx, PyYAML.
"""
import importlib.util
import os
import subprocess
import sys
import tempfile

import yaml
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

HERE = os.path.dirname(os.path.abspath(__file__))

# Font fallback (theme fonts + font table with panose/alt names) is shared
# with the document template, so a machine without Century Gothic falls back
# to a sans (URW Gothic), not a serif.
sys.dont_write_bytecode = True
_spec = importlib.util.spec_from_file_location(
    "dg_reference", os.path.join(HERE, "..", "pandoc-docx", "make-reference.py"))
DOCREF = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(DOCREF)

# Brand rules (Marketing, 2026-11-22): Arial only, #303030 and #E21E3A.
RED, INK, SLATE, MUTED, LINE = "E21E3A", "303030", "303030", "303030", "D9D9D9"
HEAD_FONT, BODY_FONT = "Century Gothic", "Arial"   # as ../pandoc-docx
# the one PNG rendering of the registered logo in the repository
LOGO = os.path.join(HERE, "..", "pandoc-docx", "daignostics-logo.png")
VARIANTS = ("a", "b", "c")

LABELS = {
    "hr": dict(page="Stranica", court="Upisano kod", mbs="MBS", oib="OIB",
               capital="Temeljni kapital", management="Direktor", bank="Banka"),
    "en": dict(page="Page", court="Registered at", mbs="Reg. no. (MBS)",
               oib="VAT ID", capital="Share capital",
               management="Managing director", bank="Bank"),
}


def rgb(h):
    return RGBColor.from_string(h)


def set_font(run_or_style_font, name, size=None, color=None, bold=None):
    f = run_or_style_font
    f.name = name
    if size:
        f.size = Pt(size)
    if color:
        f.color.rgb = rgb(color)
    if bold is not None:
        f.bold = bold


def east_asia(style_or_run, name):
    rpr = style_or_run.element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rfonts.set(qn(a), name)


PPR_AFTER_PBDR = ("w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
                  "w:wordWrap", "w:overflowPunct", "w:topLinePunct",
                  "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi",
                  "w:adjustRightInd", "w:snapToGrid", "w:spacing", "w:ind",
                  "w:contextualSpacing", "w:mirrorIndents",
                  "w:suppressOverlap", "w:jc", "w:textDirection",
                  "w:textAlignment", "w:textboxTightWrap",
                  "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
                  "w:sectPr", "w:pPrChange")


def insert_before(parent, el, *succ):
    """Insert el before the first existing child whose tag is in succ."""
    tags = {qn(s) for s in succ}
    for i, child in enumerate(parent):
        if child.tag in tags:
            parent.insert(i, el)
            return el
    parent.append(el)
    return el


def para_border(p, side, color, sz=6, space=4):
    """Add a paragraph border, keeping the CT_PPr child order valid."""
    ppr = p._p.get_or_add_pPr() if hasattr(p, "_p") else p.element.get_or_add_pPr()
    bdr = ppr.find(qn("w:pBdr"))
    if bdr is None:
        bdr = OxmlElement("w:pBdr")
        insert_before(ppr, bdr, *PPR_AFTER_PBDR)
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    order = ["top", "left", "bottom", "right", "between", "bar"]
    succ = [f"w:{x}" for x in order[order.index(side) + 1:]]
    insert_before(bdr, el, *succ)


def style(doc, name, base="Normal", size=10.5, color=INK, font=BODY_FONT,
          bold=False, before=0, after=0, keep_next=False, line=None):
    st = doc.styles
    s = None
    for x in st:
        if x.name and x.name.lower() == name.lower():
            s = x
            break
    if s is None:
        s = st.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    if base:
        s.base_style = [x for x in st if x.name == base][0]
    set_font(s.font, font, size, color, bold)
    east_asia(s, font)
    pf = s.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.keep_with_next = keep_next
    if line:
        pf.line_spacing = line
    return s


def field(par, instr, size=None, color=None):
    """Append a simple field (PAGE, NUMPAGES) to a paragraph."""
    r = par.add_run()
    if size:
        r.font.size = Pt(size)
    if color:
        r.font.color.rgb = rgb(color)
    for kind, text in (("begin", None), ("instr", instr), ("separate", None),
                       ("text", "1"), ("end", None)):
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = f" {text} "
        elif kind == "text":
            el = OxmlElement("w:t")
            el.text = text
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
        r._r.append(el)
    return r


def wordmark(par, height_mm):
    """Official dAIgnostics logo (PNG rendered from the vector SVG)."""
    par.add_run().add_picture(LOGO, height=Mm(height_mm))


def slogan(par, size):
    for txt, col in (("Detect earlier. ", INK), ("Act smarter.", RED),
                     (" For a life worth living.", INK)):
        r = par.add_run(txt)
        set_font(r.font, BODY_FONT, size, col, True)
        east_asia(r, BODY_FONT)


def no_borders(table):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:val"), "nil")
        b.append(e)
    insert_before(tblPr, b, "w:shd", "w:tblLayout", "w:tblCellMar",
                                "w:tblLook", "w:tblCaption",
                                "w:tblDescription", "w:tblPrChange")
    mar = OxmlElement("w:tblCellMar")
    for side in ("left", "right"):
        e = OxmlElement(f"w:{side}")
        e.set(qn("w:w"), "0")
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    insert_before(tblPr, mar, "w:tblLook", "w:tblCaption",
                                "w:tblDescription", "w:tblPrChange")


def cell_text(cell, lines, size, color, first_bold=False, align=None):
    cell.paragraphs[0].text = ""
    for i, line in enumerate(lines):
        p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.0
        if align:
            p.alignment = align
        r = p.add_run(line)
        set_font(r.font, BODY_FONT, size,
                 SLATE if (first_bold and i == 0) else color,
                 bool(first_bold and i == 0))
        east_asia(r, BODY_FONT)


def legal_footer(footer, c, lang, width):
    L = LABELS[lang]
    for p in footer.paragraphs:
        p.text = ""
    p0 = footer.paragraphs[0]
    p0.paragraph_format.space_after = Pt(0)
    para_border(p0, "bottom", LINE, sz=4, space=1)
    country = f", {c['company-country-en']}" if lang == "en" else ""
    capital = c["company-capital-en"] if lang == "en" else c["company-capital"]
    paid = c.get("company-capital-paid-en" if lang == "en" else "company-capital-paid")
    if paid:
        capital += f", {paid}"
    court = c["company-court-en"] if lang == "en" else c["company-court"]
    oib = c["company-vat"] if lang == "en" else c["company-oib"]
    cols = [
        [c["company-name"], f"{c['company-street']}, {c['company-city']}{country}",
         f"{c['company-web']} · {c['company-email']}", c["company-phone"]],
        [f"{L['court']}: {court}", f"{L['mbs']}: {c['company-mbs']} · {L['oib']}: {oib}",
         f"{L['capital']}: {capital}", f"{L['management']}: {c['company-management']}"],
        [f"{L['bank']}: {c['company-bank']}", f"IBAN: {c['company-iban']}",
         f"SWIFT/BIC: {c['company-swift']}"],
    ]
    t = footer.add_table(rows=1, cols=3, width=width)
    no_borders(t)
    widths = [Mm(58), Mm(57), Mm(50)]
    for i, lines in enumerate(cols):
        t.cell(0, i).width = widths[i]
        cell_text(t.cell(0, i), lines, 6.5, MUTED, first_bold=(i == 0))


def contact_lines(company, lang):
    """Address + T: / E: / W: with bold labels (brand rule), data regular."""
    country = f", {company['company-country-en']}" if lang == "en" else ""
    items = [(None, f"{company['company-street']}, {company['company-city']}{country}"),
             ("T:", company["company-phone"]), ("E:", company["company-email"]),
             ("W:", "www." + company["company-web"])]
    return items


def add_runs(p, items, size, sep=None):
    for i, (lab, val) in enumerate(items):
        if i and sep is not None:
            r = p.add_run(sep)
            set_font(r.font, BODY_FONT, size, INK, False)
        if lab:
            r = p.add_run(lab + " ")
            set_font(r.font, BODY_FONT, size, INK, True)
            east_asia(r, BODY_FONT)
        r = p.add_run(val)
        set_font(r.font, BODY_FONT, size, INK, False)
        east_asia(r, BODY_FONT)


def tight(p, before=0, after=0):
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.line_spacing = 1.0


def shrink_first(fh):
    # python-docx leaves pandoc's empty first paragraph above a header
    # table: shrink it to (almost) nothing.
    p0 = fh.paragraphs[0]
    p0.paragraph_format.space_after = Pt(0)
    p0.paragraph_format.line_spacing = Pt(1)
    p0.add_run("").font.size = Pt(1)


def cell_border(cell, side, color, sz):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    e = OxmlElement(f"w:{side}")
    e.set(qn("w:val"), "single")
    e.set(qn("w:sz"), str(sz))
    e.set(qn("w:space"), "0")
    e.set(qn("w:color"), color)
    b.append(e)
    insert_before(tcPr, b, "w:shd", "w:noWrap", "w:tcMar", "w:textDirection",
                  "w:tcFitText", "w:vAlign", "w:hideMark")


def cell_margin_left(cell, mm):
    tcPr = cell._tc.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    e = OxmlElement("w:left")
    e.set(qn("w:w"), str(int(mm * 56.7)))
    e.set(qn("w:type"), "dxa")
    mar.append(e)
    insert_before(tcPr, mar, "w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")


def letterhead_a(fh, company, lang, width):
    """A: logo + slogan left, contact right-aligned, red rule under both."""
    t = fh.add_table(rows=1, cols=2, width=width)
    no_borders(t)
    t.cell(0, 0).width, t.cell(0, 1).width = Mm(95), Mm(70)
    p = t.cell(0, 0).paragraphs[0]
    p.paragraph_format.space_before = Mm(1)
    p.paragraph_format.space_after = Pt(2)
    wordmark(p, 11)
    sp = t.cell(0, 0).add_paragraph()
    sp.paragraph_format.space_after = Pt(0)
    slogan(sp, 7.5)
    items = contact_lines(company, lang)
    cell = t.cell(0, 1)
    for i, it in enumerate(items):
        cp = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        tight(cp)
        cp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        add_runs(cp, [it], 7.5)
    shrink_first(fh)
    rule = fh.add_paragraph()
    rule.paragraph_format.space_before = Pt(2)
    rule.paragraph_format.space_after = Pt(0)
    para_border(rule, "bottom", RED, sz=8, space=1)


def letterhead_b(fh, company, lang, width):
    """B: centred logo, slogan, short red rule, contact in one line."""
    p = fh.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(p, after=3)
    wordmark(p, 12)
    sp = fh.add_paragraph()
    sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(sp, after=2)
    slogan(sp, 7.5)
    rule = fh.add_paragraph()
    tight(rule)
    rule.paragraph_format.left_indent = rule.paragraph_format.right_indent = Mm(70)
    rule.paragraph_format.line_spacing = Pt(4)
    para_border(rule, "bottom", RED, sz=8, space=1)
    cp = fh.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tight(cp, before=4)
    add_runs(cp, contact_lines(company, lang), 7.2, sep="\u2003")


def letterhead_c(fh, company, lang, width):
    """C: logo + slogan | red vertical line | contact (as the e-mail signature)."""
    t = fh.add_table(rows=1, cols=2, width=width)
    no_borders(t)
    t.cell(0, 0).width, t.cell(0, 1).width = Mm(107), Mm(58)
    p = t.cell(0, 0).paragraphs[0]
    p.paragraph_format.space_before = Mm(1)
    p.paragraph_format.space_after = Pt(2)
    wordmark(p, 10)
    sp = t.cell(0, 0).add_paragraph()
    sp.paragraph_format.space_after = Pt(0)
    slogan(sp, 7.5)
    cell = t.cell(0, 1)
    cell_border(cell, "left", RED, 12)
    cell_margin_left(cell, 4)
    for i, it in enumerate(contact_lines(company, lang)):
        cp = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
        tight(cp)
        cp.paragraph_format.line_spacing = Pt(10)
        add_runs(cp, [it], 7.5)
    shrink_first(fh)
    # same vertical room as A's rule paragraph keeps the address window put
    # (measured: without it the address sits 5.7 mm higher than in A / B)
    pad = fh.add_paragraph()
    pad.paragraph_format.space_before = Mm(6.4)
    pad.paragraph_format.space_after = Pt(0)
    pad.paragraph_format.line_spacing = Pt(2)


LETTERHEADS = {"a": letterhead_a, "b": letterhead_b, "c": letterhead_c}


def build(lang, company, variant):
    tmp = tempfile.mkdtemp()
    base = os.path.join(tmp, "ref.docx")
    with open(base, "wb") as fh:
        fh.write(subprocess.check_output(
            ["pandoc", "--print-default-data-file", "reference.docx"]))
    doc = Document(base)

    # ---- page setup -------------------------------------------------------
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin, sec.right_margin = Mm(25), Mm(20)
    sec.top_margin, sec.bottom_margin = Mm(27), Mm(34)
    sec.header_distance, sec.footer_distance = Mm(12), Mm(8)
    sec.different_first_page_header_footer = True
    width = Mm(165)

    # ---- styles -----------------------------------------------------------
    style(doc, "Normal", base=None, size=10.5, line=1.1)
    for n in ("Body Text", "First Paragraph", "Compact"):
        try:
            style(doc, n, size=10.5, before=0, after=8, line=1.1)
        except KeyError:
            pass
    style(doc, "Address Spacer", size=2, before=Pt(0).pt)
    doc.styles["Address Spacer"].paragraph_format.space_before = Mm(11)
    s = style(doc, "Return Address", size=6.5, color=MUTED, after=3)
    s.paragraph_format.space_before = Mm(9)
    para_border(s, "bottom", LINE, sz=4, space=2)
    s.paragraph_format.right_indent = Mm(8)
    style(doc, "Recipient", size=10.5, after=0, line=1.15)
    style(doc, "Reference Label", size=7.5, color=MUTED, before=3)
    style(doc, "Reference Value", size=8, color=INK)
    s = style(doc, "Letter Subject", size=12, font=HEAD_FONT, bold=True,
              color=INK, before=10, after=10, keep_next=True)
    style(doc, "Closing", size=10.5, before=6, after=0, keep_next=True)
    s = style(doc, "Signature Space", size=10.5, keep_next=True)
    s.paragraph_format.space_before = Mm(10)
    style(doc, "Signature Image", size=10.5, before=2, after=0, keep_next=True)
    style(doc, "Signatory", size=10.5, bold=True, keep_next=True)
    style(doc, "Signatory Title", size=10.5, bold=True, color=RED, keep_next=True)
    style(doc, "Enclosure Label", size=9, bold=True, before=10, keep_next=True)
    style(doc, "Enclosure Item", size=9)
    style(doc, "Signatory Company", size=10.5)
    for hs in ("Header", "Footer"):
        try:
            style(doc, hs, size=8, color=MUTED)
        except Exception:
            pass
    style(doc, "Funding Note", size=6.5, color=MUTED, before=8)
    for h in ("Heading 1", "Heading 2", "Heading 3"):
        style(doc, h, size={"Heading 1": 12, "Heading 2": 11, "Heading 3": 10.5}[h],
              font=HEAD_FONT, bold=True, color=INK, before=10, after=4, keep_next=True)

    # ---- first-page letterhead (variant A, B or C) -------------------------
    fh = sec.first_page_header
    fh.is_linked_to_previous = False
    LETTERHEADS[variant](fh, company, lang, width)

    # ---- running header (pages 2+): wordmark ......... Stranica X / Y -----
    h = sec.header
    h.is_linked_to_previous = False
    hp = h.paragraphs[0]
    hp.text = ""
    hp.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
    para_border(hp, "bottom", RED, sz=6, space=4)
    wordmark(hp, 4.2)
    r = hp.add_run(f"\t{LABELS[lang]['page']} ")
    set_font(r.font, BODY_FONT, 8, MUTED)
    field(hp, "PAGE", 8, MUTED)
    r = hp.add_run(" / ")
    set_font(r.font, BODY_FONT, 8, MUTED)
    field(hp, "NUMPAGES", 8, MUTED)

    # ---- footers: statutory details on every page ---------------------------
    for ft in (sec.first_page_footer, sec.footer):
        ft.is_linked_to_previous = False
        legal_footer(ft, company, lang, width)

    # ---- remove pandoc's sample body ----------------------------------------
    body = doc.element.body
    for el in list(body):
        if el.tag != qn("w:sectPr"):
            body.remove(el)

    DOCREF.patch_part_blob(doc, "/word/theme/theme1.xml", DOCREF.patch_theme)
    DOCREF.patch_part_blob(doc, "/word/fontTable.xml", DOCREF.patch_font_table)

    out = os.path.join(HERE, f"reference-letter-{lang}-{variant}.docx")
    doc.save(out)
    print("Wrote", out)


if __name__ == "__main__":
    with open(os.path.join(HERE, "company.yaml"), encoding="utf-8") as fh:
        company = yaml.safe_load(fh)
    for lang in ("hr", "en"):
        for variant in VARIANTS:
            build(lang, company, variant)
