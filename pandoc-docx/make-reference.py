#!/usr/bin/env python3
"""Generate reference.docx (and daignostics.theme) for the dAIgnostics DOCX template.

    python3 make-reference.py                      # writes reference.docx + daignostics.theme
    python3 make-reference.py --page-label "Page {page} of {total}"

Starts from pandoc's own default reference.docx (so every style pandoc emits
exists), then restyles it to match ../pandoc-document/daignostics.latex:
palette, fonts, headings, callout and code styles, booktabs-like table style,
the cover-page styles used by cover-docx.lua, A4 page geometry and the
running header/footer (logo, title, project, page X / Y, EU emblem).

Requires: pandoc on PATH, python-docx (tested with 1.2.0).
"""
import argparse
import os
import re
import subprocess
import sys
import tempfile

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Twips

HERE = os.path.dirname(os.path.abspath(__file__))
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSDECL = f'xmlns:w="{W_NS}"'

# ---- Brand palette (same names/values as daignostics.latex) -------------
RED = "E21E3A"    # dgred   - accent (logo red)
INK = "303030"    # dgink   - headings / text (logo ink)
SLATE = "2B3A4A"  # dgslate - secondary dark
MUTED = "5A6B7B"  # dgmuted - captions / running heads
LINE = "D7DEE4"   # dgline  - hairlines
BG = "F4F6F8"     # dgbg    - code / callout fill
EU = "003399"     # dgeu    - EU blue (tip/info spine)

# ---- Fonts --------------------------------------------------------------
# Word machines do not have TeX Gyre fonts. Century Gothic ships with
# Windows/Office and is what TeX Gyre Adventor imitates; Arial is the Office
# stand-in for Helvetica (= TeX Gyre Heros). Styles name the font explicitly
# AND point at the theme major/minor fonts, so if Century Gothic is missing
# Word/LibreOffice fall back through the theme + font-table panose data to a
# plain sans rather than to a serif.
HEAD_FONT = "Century Gothic"
BODY_FONT = "Arial"
MONO_FONT = "Consolas"

# ---- Page geometry (A4, same margins as the LaTeX template) -------------
PAGE_W, PAGE_H = 11906, 16838          # twips
MARGIN_LR = 1361                        # 2.4 cm
MARGIN_T, MARGIN_B = 1474, 1531         # 2.6 cm / 2.7 cm
HEADER_D, FOOTER_D = 680, 397           # 1.2 cm / 0.7 cm
TEXT_W = PAGE_W - 2 * MARGIN_LR         # 9184 twips = 16.2 cm


def head_fonts():
    return (f'<w:rFonts w:ascii="{HEAD_FONT}" w:hAnsi="{HEAD_FONT}" '
            f'w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" '
            f'w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/>')


def body_fonts():
    return (f'<w:rFonts w:ascii="{BODY_FONT}" w:hAnsi="{BODY_FONT}" '
            f'w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" '
            f'w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/>')


def mono_fonts():
    return f'<w:rFonts w:ascii="{MONO_FONT}" w:hAnsi="{MONO_FONT}" w:cs="{MONO_FONT}"/>'


def sz(pt):
    hp = int(round(pt * 2))
    return f'<w:sz w:val="{hp}"/><w:szCs w:val="{hp}"/>'


def pstyle(sid, name, ppr="", rpr="", based="Normal", nxt=None, custom=True,
           extra=""):
    attrs = ' w:customStyle="1"' if custom else ""
    nx = f'<w:next w:val="{nxt}"/>' if nxt else ""
    bo = f'<w:basedOn w:val="{based}"/>' if based else ""
    return (f'<w:style {NSDECL} w:type="paragraph"{attrs} w:styleId="{sid}">'
            f'<w:name w:val="{name}"/>{bo}{nx}{extra}<w:qFormat/>'
            f'<w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>')


def cstyle(sid, name, rpr, based="DefaultParagraphFont", custom=True):
    attrs = ' w:customStyle="1"' if custom else ""
    return (f'<w:style {NSDECL} w:type="character"{attrs} w:styleId="{sid}">'
            f'<w:name w:val="{name}"/><w:basedOn w:val="{based}"/>'
            f'<w:rPr>{rpr}</w:rPr></w:style>')


def box_ppr(spine, spine_sz=24, before=0, after=0, keep_next=False, wrap_off=False):
    """Paragraph props for a shaded box with a coloured left spine.

    Consecutive paragraphs with identical borders/indents are merged by Word
    and LibreOffice into one box; the top/bottom "borders" are drawn in the
    fill colour and only act as inner padding of the merged group."""
    kn = "<w:keepNext/>" if keep_next else ""
    return (f'{kn}<w:pBdr>'
            f'<w:top w:val="single" w:sz="4" w:space="6" w:color="{BG}"/>'
            f'<w:left w:val="single" w:sz="{spine_sz}" w:space="9" w:color="{spine}"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="6" w:color="{BG}"/>'
            f'<w:right w:val="single" w:sz="4" w:space="9" w:color="{BG}"/>'
            f'</w:pBdr><w:shd w:val="clear" w:color="auto" w:fill="{BG}"/>'
            + ('<w:wordWrap w:val="off"/>' if wrap_off else '') +
            f'<w:spacing w:before="{before}" w:after="{after}" w:line="264" w:lineRule="auto"/>'
            f'<w:ind w:left="250" w:right="190"/><w:jc w:val="left"/>')


def build_styles():
    S = []
    body_ppr = '<w:spacing w:before="0" w:after="140" w:line="269" w:lineRule="auto"/>'
    # Normal / body
    S.append(pstyle("Normal", "Normal", body_ppr.replace('w:after="140"', 'w:after="0"'),
                    body_fonts() + f'<w:color w:val="{INK}"/>' + sz(10),
                    based=None, custom=False))
    S.append(pstyle("BodyText", "Body Text", body_ppr + '<w:jc w:val="both"/>', "",
                    custom=False, extra='<w:link w:val="BodyTextChar"/>'))
    S.append(pstyle("FirstParagraph", "First Paragraph", "", "", based="BodyText",
                    nxt="BodyText"))
    S.append(pstyle("Compact", "Compact",
                    '<w:spacing w:before="30" w:after="30"/><w:jc w:val="left"/>', "",
                    based="BodyText"))
    S.append(pstyle("BlockText", "Block Text",
                    f'<w:pBdr><w:left w:val="single" w:sz="12" w:space="8" w:color="{LINE}"/></w:pBdr>'
                    '<w:spacing w:before="100" w:after="100"/><w:ind w:left="300" w:right="300"/>',
                    f'<w:color w:val="{SLATE}"/>', based="BodyText", nxt="BodyText", custom=False))

    # Pandoc's own title block is replaced by the cover built in cover-docx.lua.
    # The paragraphs stay in the file (they feed nothing visible) so they are
    # hidden; title/author still populate the document properties.
    hidden = '<w:vanish/>' + sz(1)
    tiny = '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'
    for sid, name, builtin in (("Title", "Title", False), ("Subtitle", "Subtitle", False),
                               ("Author", "Author", True), ("Date", "Date", False),
                               ("AbstractTitle", "Abstract Title", True)):
        S.append(pstyle(sid, name, tiny, hidden, custom=builtin, nxt="BodyText"))
    S.append(pstyle("Abstract", "Abstract",
                    '<w:spacing w:before="100" w:after="300"/><w:ind w:left="567" w:right="567"/>',
                    f'<w:i/><w:color w:val="{SLATE}"/>', based="BodyText"))

    # Headings: geometric font, slate ink, red section number (SectionNumber),
    # hairline under level 1 (as \titlerule in the LaTeX template).
    def heading(n, size, color, before, after, rule=False):
        bdr = (f'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="3" w:color="{LINE}"/></w:pBdr>'
               if rule else "")
        ppr = (f'<w:keepNext/><w:keepLines/>{bdr}'
               f'<w:spacing w:before="{before}" w:after="{after}" w:line="240" w:lineRule="auto"/>'
               f'<w:jc w:val="left"/><w:outlineLvl w:val="{n-1}"/>')
        rpr = head_fonts() + f'<w:b/><w:bCs/><w:color w:val="{color}"/>' + sz(size)
        return pstyle(f"Heading{n}", f"heading {n}", ppr, rpr, nxt="BodyText", custom=False,
                      extra=f'<w:uiPriority w:val="9"/>')
    S.append(heading(1, 17, INK, 440, 180, rule=True))
    S.append(heading(2, 13, INK, 320, 110))
    S.append(heading(3, 12, SLATE, 260, 90))
    S.append(heading(4, 10.5, SLATE, 200, 60))
    S.append(heading(5, 10, SLATE, 160, 40))
    S.append(heading(6, 10, MUTED, 160, 40))
    S.append(cstyle("SectionNumber", "Section Number", f'<w:color w:val="{RED}"/>'))

    # TOC
    S.append(pstyle("TOCHeading", "TOC Heading",
                    f'<w:keepNext/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="3" w:color="{LINE}"/></w:pBdr>'
                    '<w:spacing w:before="0" w:after="200"/><w:jc w:val="left"/><w:outlineLvl w:val="9"/>',
                    head_fonts() + f'<w:b/><w:color w:val="{INK}"/>' + sz(18),
                    nxt="BodyText", custom=False))
    S.append(pstyle("TOC1", "toc 1",
                    f'<w:tabs><w:tab w:val="left" w:pos="400"/><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs>'
                    '<w:spacing w:before="140" w:after="20"/>',
                    '<w:b/>', nxt="BodyText", custom=False))
    S.append(pstyle("TOC2", "toc 2",
                    f'<w:tabs><w:tab w:val="left" w:pos="900"/><w:tab w:val="right" w:leader="dot" w:pos="{TEXT_W}"/></w:tabs>'
                    '<w:spacing w:before="0" w:after="20"/><w:ind w:left="400"/>',
                    "", nxt="BodyText", custom=False))
    S.append(pstyle("TOC3", "toc 3",
                    f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{TEXT_W}"/></w:tabs>'
                    '<w:spacing w:before="0" w:after="20"/><w:ind w:left="900"/>',
                    "", nxt="BodyText", custom=False))

    # Code
    S.append(pstyle("SourceCode", "Source Code",
                    box_ppr(RED, spine_sz=20, before=160, after=180, wrap_off=True)
                    .replace('w:line="264"', 'w:line="252"'),
                    mono_fonts() + f'<w:color w:val="{INK}"/>' + sz(8.5),
                    extra='<w:link w:val="VerbatimChar"/>'))
    S.append(cstyle("VerbatimChar", "Verbatim Char", mono_fonts() + sz(9),
                    based="BodyTextChar"))
    S.append(cstyle("BodyTextChar", "Body Text Char", "", custom=True))
    S.append(cstyle("Hyperlink", "Hyperlink", f'<w:color w:val="{RED}"/>',
                    based="BodyTextChar", custom=False))

    # Captions, figures, footnotes, definitions
    S.append(pstyle("Caption", "caption", '<w:spacing w:before="60" w:after="200"/><w:jc w:val="center"/>',
                    f'<w:i/><w:color w:val="{MUTED}"/>' + sz(9), custom=False))
    # pandoc puts table captions above the table: keep them on its page
    S.append(pstyle("TableCaption", "Table Caption", "<w:keepNext/>", "", based="Caption"))
    S.append(pstyle("ImageCaption", "Image Caption", "", "", based="Caption"))
    S.append(pstyle("FootnoteText", "footnote text", '<w:spacing w:after="40"/>', sz(8.5),
                    custom=False))

    # Callouts (admonitions-docx.lua): <Kind>Title + <Kind> body
    for kind, color in (("Note", SLATE), ("Tip", EU), ("Warning", RED)):
        S.append(pstyle(f"Callout{kind}Title", f"Callout {kind} Title",
                        box_ppr(color, before=180, after=70, keep_next=True),
                        head_fonts() + f'<w:b/><w:bCs/><w:color w:val="{color}"/>' + sz(9.5),
                        nxt=f"Callout{kind}"))
        S.append(pstyle(f"Callout{kind}", f"Callout {kind}",
                        box_ppr(color, before=0, after=70), "", nxt="BodyText"))
    # empty separator after each callout: keeps two adjacent callouts of the
    # same kind from merging into one box
    S.append(pstyle("CalloutEnd", "Callout End",
                    '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>',
                    sz(2), nxt="BodyText"))

    # Cover page (cover-docx.lua)
    S.append(pstyle("CoverBrand", "Cover Brand",
                    f'<w:pBdr><w:bottom w:val="single" w:sz="24" w:space="10" w:color="{RED}"/></w:pBdr>'
                    f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs>'
                    '<w:spacing w:before="280" w:after="0"/><w:jc w:val="left"/>',
                    head_fonts() + f'<w:color w:val="{MUTED}"/>' + sz(11)))
    S.append(pstyle("CoverDoctype", "Cover Doctype",
                    '<w:keepNext/><w:spacing w:before="3700" w:after="260"/><w:jc w:val="left"/>',
                    head_fonts() + f'<w:b/><w:caps/><w:color w:val="{RED}"/>' + sz(12.5),
                    nxt="CoverTitle"))
    S.append(pstyle("CoverTitle", "Cover Title",
                    '<w:keepNext/><w:spacing w:before="0" w:after="0" w:line="216" w:lineRule="auto"/><w:jc w:val="left"/>',
                    head_fonts() + f'<w:b/><w:color w:val="{INK}"/>' + sz(32),
                    nxt="CoverSubtitle"))
    S.append(pstyle("CoverSubtitle", "Cover Subtitle",
                    '<w:keepNext/><w:spacing w:before="240" w:after="0"/><w:jc w:val="left"/>',
                    head_fonts() + f'<w:color w:val="{MUTED}"/>' + sz(16)))
    S.append(pstyle("CoverRule", "Cover Rule",
                    f'<w:pBdr><w:bottom w:val="single" w:sz="10" w:space="1" w:color="{LINE}"/></w:pBdr>'
                    f'<w:spacing w:before="0" w:after="400" w:line="400" w:lineRule="exact"/>'
                    f'<w:ind w:right="{TEXT_W - 3402}"/>', sz(2)))
    S.append(pstyle("CoverMeta", "Cover Meta",
                    '<w:tabs><w:tab w:val="left" w:pos="1250"/></w:tabs>'
                    '<w:spacing w:before="0" w:after="70"/><w:ind w:left="1250" w:hanging="1250"/><w:jc w:val="left"/>',
                    head_fonts() + f'<w:color w:val="{INK}"/>' + sz(10.5)))
    S.append(cstyle("CoverLabel", "Cover Label", f'<w:b/><w:color w:val="{SLATE}"/>'))
    S.append(pstyle("CoverBreak", "Cover Break",
                    '<w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>', sz(1)))
    S.append(cstyle("Wordmark", "Wordmark", head_fonts() + f'<w:b/><w:color w:val="{INK}"/>' + sz(12)))
    S.append(cstyle("BrandAccent", "Brand Accent", f'<w:color w:val="{RED}"/>', based="Wordmark"))

    # Running header / footer paragraphs
    S.append(pstyle("Header", "header",
                    f'<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="6" w:color="{RED}"/></w:pBdr>'
                    f'<w:tabs><w:tab w:val="right" w:pos="{TEXT_W}"/></w:tabs><w:spacing w:after="0"/>',
                    sz(8), custom=False))
    S.append(pstyle("Footer", "footer", '<w:spacing w:after="0"/>',
                    f'<w:color w:val="{MUTED}"/>' + sz(8), custom=False))
    return S


def table_style():
    b = lambda side, size: f'<w:{side} w:val="single" w:sz="{size}" w:space="0" w:color="{INK}"/>'
    return (f'<w:style {NSDECL} w:type="table" w:default="1" w:styleId="Table">'
            '<w:name w:val="Table"/><w:basedOn w:val="TableNormal"/><w:qFormat/>'
            '<w:pPr><w:spacing w:before="0" w:after="0"/></w:pPr>'
            f'<w:rPr>{sz(9.5)}</w:rPr>'
            '<w:tblPr><w:jc w:val="center"/><w:tblInd w:w="0" w:type="dxa"/>'
            f'<w:tblBorders>{b("top", 8)}{b("bottom", 8)}</w:tblBorders>'
            '<w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:left w:w="120" w:type="dxa"/>'
            '<w:bottom w:w="30" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>'
            '</w:tblPr>'
            f'<w:tblStylePr w:type="firstRow"><w:rPr><w:color w:val="{INK}"/></w:rPr>'
            f'<w:tcPr><w:tcBorders>{b("bottom", 4)}</w:tcBorders><w:vAlign w:val="bottom"/></w:tcPr>'
            '</w:tblStylePr></w:style>')


def replace_styles(doc, xml_list):
    styles = doc.styles.element
    for xml in xml_list:
        el = parse_xml(xml)
        sid = el.get(qn("w:styleId"))
        for old in styles.findall(qn("w:style")):
            if old.get(qn("w:styleId")) == sid:
                # keep pandoc's link/ui metadata? simpler: replace wholesale
                styles.remove(old)
        styles.append(el)


def set_doc_defaults(doc):
    styles = doc.styles.element
    dd = styles.find(qn("w:docDefaults"))
    new = parse_xml(
        f'<w:docDefaults {NSDECL}><w:rPrDefault><w:rPr>{body_fonts()}{sz(10)}'
        '<w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="269" w:lineRule="auto"/></w:pPr>'
        '</w:pPrDefault></w:docDefaults>')
    dd.addprevious(new)
    styles.remove(dd)


def patch_part_blob(doc, partname, fn):
    for part in doc.part.package.iter_parts():
        if str(part.partname) == partname:
            if hasattr(part, "_element"):      # XmlPart
                xml = part._element.xml
                part._element = parse_xml(fn(xml).encode("utf-8"))
            else:
                part._blob = fn(part.blob.decode("utf-8")).encode("utf-8")
            return
    raise SystemExit(f"part {partname} not found")


def patch_theme(xml):
    xml = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*(" panose=")[^"]*"',
                 r'\g<1>' + HEAD_FONT + r'\g<2>020B0502020202020204"', xml)
    xml = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*(" panose=")[^"]*"',
                 r'\g<1>' + BODY_FONT + r'\g<2>020B0604020202020204"', xml)
    xml = re.sub(r'<a:fontScheme name="[^"]*"', '<a:fontScheme name="dAIgnostics"', xml)
    return xml


FONT_TABLE_ENTRIES = {
    HEAD_FONT: ('<w:altName w:val="URW Gothic"/><w:panose1 w:val="020B0502020202020204"/>'
                '<w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/>'),
    BODY_FONT: ('<w:altName w:val="Helvetica"/><w:panose1 w:val="020B0604020202020204"/>'
                '<w:charset w:val="00"/><w:family w:val="swiss"/><w:pitch w:val="variable"/>'),
    MONO_FONT: ('<w:panose1 w:val="020B0609020204030204"/><w:charset w:val="00"/>'
                '<w:family w:val="modern"/><w:pitch w:val="fixed"/>'),
}


def patch_font_table(xml):
    for name, body in FONT_TABLE_ENTRIES.items():
        if f'w:name="{name}"' not in xml:
            xml = xml.replace("</w:fonts>", f'<w:font w:name="{name}">{body}</w:font></w:fonts>')
    return xml


def field_runs(instr, rpr="", cached=""):
    """Simple field as runs; w:dirty asks Word to refresh it on open."""
    r = lambda inner: f'<w:r {NSDECL}><w:rPr>{rpr}</w:rPr>{inner}</w:r>'
    return [
        r('<w:fldChar w:fldCharType="begin" w:dirty="true"/>'),
        r(f'<w:instrText xml:space="preserve"> {instr} </w:instrText>'),
        r('<w:fldChar w:fldCharType="separate"/>'),
        r(f'<w:t xml:space="preserve">{cached}</w:t>'),
        r('<w:fldChar w:fldCharType="end"/>'),
    ]


def text_run(text, rpr=""):
    return f'<w:r {NSDECL}><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def append_xml(parent, xml):
    parent.append(parse_xml(xml))


def footer_table(doc, part, fractions, valign="center"):
    """Borderless full-width one-row table in a header/footer part."""
    widths = [int(TEXT_W * f) for f in fractions[:-1]]
    widths.append(TEXT_W - sum(widths))
    tbl = part.add_table(1, len(widths), Twips(TEXT_W))
    tblPr = tbl._tbl.tblPr
    for child in list(tblPr):
        tblPr.remove(child)
    # schema order; TableNormal so pandoc's default "Table" style does not leak in
    for xml in (f'<w:tblStyle {NSDECL} w:val="TableNormal"/>',
                f'<w:tblW {NSDECL} w:w="{TEXT_W}" w:type="dxa"/>',
                f'<w:tblBorders {NSDECL}><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/>'
                '<w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>',
                f'<w:tblLayout {NSDECL} w:type="fixed"/>',
                f'<w:tblCellMar {NSDECL}><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar>',
                f'<w:tblLook {NSDECL} w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" '
                'w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>'):
        tblPr.append(parse_xml(xml))
    for gc, w in zip(tbl._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(w))
    cells = tbl.rows[0].cells
    for i, (cell, w) in enumerate(zip(cells, widths)):
        cell.width = Twips(w)
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:vAlign {NSDECL} w:val="{valign}"/>'))
        jc = "left" if i == 0 else ("right" if i == len(cells) - 1 else "center")
        cp = cell.paragraphs[0]
        cp.style = doc.styles["Footer"]
        cp._p.get_or_add_pPr().append(parse_xml(f'<w:jc {NSDECL} w:val="{jc}"/>'))
    # the empty paragraph Word requires after a table: make it tiny
    last = part.paragraphs[-1]
    last.style = doc.styles["Footer"]
    last._p.get_or_add_pPr().append(parse_xml(
        f'<w:spacing {NSDECL} w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'))
    return [c.paragraphs[0] for c in cells]


def build_header_footer(doc, logo, brand_logo, page_label):
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Twips(PAGE_W), Twips(PAGE_H)
    sec.left_margin = sec.right_margin = Twips(MARGIN_LR)
    sec.top_margin, sec.bottom_margin = Twips(MARGIN_T), Twips(MARGIN_B)
    sec.header_distance, sec.footer_distance = Twips(HEADER_D), Twips(FOOTER_D)
    # Page 1 is the cover (built by cover-docx.lua) and uses the first-page
    # header/footer: no running header, funding text + large emblem at the
    # bottom. Page numbers count the cover, so "page / total" is correct in
    # every viewer (SECTIONPAGES or {= NUMPAGES - 1} would reproduce the PDF's
    # numbering in Word but are not evaluated by LibreOffice).
    for old in sec._sectPr.findall(qn("w:pgNumType")):
        sec._sectPr.remove(old)
    sec.different_first_page_header_footer = True

    # ---- running header: logo left, document title (TITLE field) right.
    # The logo (registered mark HR Z20251009) is used as supplied - only
    # scaled by height; without the file the text wordmark is used.
    hdr = sec.header
    hdr.is_linked_to_previous = False
    p = hdr.paragraphs[0]
    p.style = doc.styles["Header"]
    if os.path.exists(brand_logo):
        p.add_run().add_picture(brand_logo, height=Cm(0.45))
    else:
        wm = '<w:rStyle w:val="Wordmark"/>'
        for xml in (text_run("d", wm),
                    text_run("AI", wm + f'<w:color w:val="{RED}"/>'),
                    text_run("gnostics", wm)):
            append_xml(p._p, xml)
    append_xml(p._p, f'<w:r {NSDECL}><w:tab/></w:r>')
    for xml in field_runs("TITLE", f'<w:i/><w:color w:val="{MUTED}"/>'):
        append_xml(p._p, xml)

    # ---- running footer: project | page / total | emblem
    ftr = sec.footer
    ftr.is_linked_to_previous = False
    left, mid, right = footer_table(doc, ftr, (0.42, 0.16, 0.42))
    # left: project or "Confidential" (custom property set by cover-docx.lua)
    for xml in field_runs('DOCPROPERTY "footer-left"'):
        append_xml(left._p, xml)
    for part in re.split(r"(\{page\}|\{total\})", page_label):
        if part == "{page}":
            xs = field_runs("PAGE", cached="1")
        elif part == "{total}":
            xs = field_runs("NUMPAGES", cached="1")
        else:
            xs = [text_run(part)] if part else []
        for xml in xs:
            append_xml(mid._p, xml)
    if os.path.exists(logo):
        right.add_run().add_picture(logo, height=Cm(1.25))

    # ---- cover (first page): empty header; funding text + 2.4 cm emblem
    fh = sec.first_page_header
    fh.is_linked_to_previous = False
    fh.paragraphs[0].style = doc.styles["Footer"]
    ff = sec.first_page_footer
    ff.is_linked_to_previous = False
    ftext, flogo = footer_table(doc, ff, (0.68, 0.32))
    for xml in field_runs('DOCPROPERTY "funding"',
                          body_fonts() + f'<w:color w:val="{MUTED}"/>' + sz(8)):
        append_xml(ftext._p, xml)
    # spacing must precede jc inside pPr
    ftext._p.get_or_add_pPr().find(qn("w:jc")).addprevious(parse_xml(
        f'<w:spacing {NSDECL} w:before="0" w:after="0" w:line="252" w:lineRule="auto"/>'))
    if os.path.exists(logo):
        flogo.add_run().add_picture(logo, height=Cm(2.4))
    # bottom padding under the cover footer, like the PDF's emblem position
    append_xml(ff._element, f'<w:p {NSDECL}><w:pPr><w:pStyle w:val="Footer"/>'
               '<w:spacing w:before="0" w:after="0" w:line="400" w:lineRule="exact"/></w:pPr></w:p>')


def write_theme(path):
    out = subprocess.run(["pandoc", "--print-highlight-style", "tango"],
                         check=True, capture_output=True, text=True).stdout
    out = out.replace('"background-color": "#f8f8f8"', f'"background-color": "#{BG.lower()}"', 1)
    with open(path, "w") as f:
        f.write(out)


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("-o", "--output", default=os.path.join(HERE, "reference.docx"))
    ap.add_argument("--logo", default=os.path.join(HERE, "eu-logo.png"),
                    help="EU emblem for the footer (default: eu-logo.png)")
    ap.add_argument("--brand-logo", default=os.path.join(HERE, "daignostics-logo.png"),
                    help="dAIgnostics logo for the running header (default: daignostics-logo.png; "
                         "text wordmark if the file is missing)")
    ap.add_argument("--page-label", default="{page} / {total}",
                    help='footer page label, e.g. "Page {page} of {total}" (default: "{page} / {total}")')
    ap.add_argument("--theme-out", default=os.path.join(HERE, "daignostics.theme"))
    a = ap.parse_args()

    with tempfile.TemporaryDirectory() as td:
        base = os.path.join(td, "base.docx")
        subprocess.run(["pandoc", "-o", base, "--print-default-data-file", "reference.docx"],
                       check=True)
        doc = Document(base)
    set_doc_defaults(doc)
    replace_styles(doc, build_styles() + [table_style()])
    patch_part_blob(doc, "/word/theme/theme1.xml", patch_theme)
    patch_part_blob(doc, "/word/fontTable.xml", patch_font_table)
    build_header_footer(doc, a.logo, a.brand_logo, a.page_label)
    cp = doc.core_properties
    cp.title, cp.author, cp.comments = "dAIgnostics reference document", "dAIgnostics", \
        "Generated by make-reference.py - pandoc --reference-doc"
    doc.save(a.output)
    write_theme(a.theme_out)
    print(f"Wrote {a.output} and {a.theme_out}")


if __name__ == "__main__":
    sys.exit(main())
