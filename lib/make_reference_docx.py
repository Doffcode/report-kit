"""Build templates/report/reference.docx from Pandoc's default reference.

Every Word style the report uses is set here, so the DOCX matches the PDF
layout rules: Times New Roman 12pt, ~1.5 spacing, chapters on a new page,
captions centred, bordered tables, A4 with a binding margin.

Run: .venv/bin/python lib/make_reference_docx.py
"""
from pathlib import Path

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt, RGBColor

KIT = Path(__file__).resolve().parent.parent
SRC = KIT / "templates/report/pandoc-default.docx"
OUT = KIT / "templates/report/reference.docx"
FONT = "Times New Roman"
BLACK = RGBColor(0, 0, 0)


def set_font(rpr_parent, name=FONT):
    """Set a concrete font on all four script slots and drop theme-font links,
    which otherwise override the name and silently bring back Calibri/Cambria."""
    rpr = rpr_parent.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for attr in list(fonts.attrib):
        if "Theme" in attr:
            del fonts.attrib[attr]
    for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(slot), name)


def style_para(d, name, size=12, bold=None, italic=None, align=None,
               before=0, after=6, spacing=1.5, page_break=False, keep_next=None, color=BLACK):
    s = d.styles[name]
    set_font(s.element)
    f = s.font
    f.size = Pt(size)
    f.color.rgb = color
    if bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    pf = s.paragraph_format
    if align is not None:
        pf.alignment = align
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = spacing
    pf.first_line_indent = Pt(0)
    pf.page_break_before = page_break
    if keep_next is not None:
        pf.keep_with_next = keep_next
    pf.widow_control = True
    return s


def table_borders(tbl_pr):
    """Single black borders on a data table (applied per table by docx_fix)."""
    old = tbl_pr.find(qn("w:tblBorders"))
    if old is not None:
        tbl_pr.remove(old)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{edge}")
        b.set(qn("w:val"), "single")
        b.set(qn("w:sz"), "4")
        b.set(qn("w:space"), "0")
        b.set(qn("w:color"), "000000")
        borders.append(b)
    tbl_pr.append(borders)


def main():
    d = docx.Document(SRC)

    # Defaults for every run in the document, so nothing falls back to theme fonts.
    defaults = d.styles.element.find(qn("w:docDefaults"))
    rpr_default = defaults.find(qn("w:rPrDefault")).find(qn("w:rPr"))
    fonts = rpr_default.find(qn("w:rFonts"))
    for attr in list(fonts.attrib):
        del fonts.attrib[attr]
    for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(slot), FONT)

    J, C, L = WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.LEFT
    for body in ("Normal", "Body Text", "First Paragraph", "Abstract", "Bibliography", "Block Text"):
        style_para(d, body, align=J)
    style_para(d, "Compact", align=J, after=2)
    style_para(d, "Footnote Text", size=10, spacing=1.0, after=2)

    # Chapters: new page, centred, uppercase is applied by the post-processor.
    style_para(d, "Heading 1", size=16, bold=True, italic=False, align=C,
               before=24, after=24, page_break=True, keep_next=True)
    style_para(d, "Heading 2", size=14, bold=True, italic=False, align=L, before=18, after=8, keep_next=True)
    style_para(d, "Heading 3", size=12, bold=True, italic=False, align=L, before=14, after=6, keep_next=True)
    style_para(d, "Heading 4", size=12, bold=True, italic=True, align=L, before=12, after=6, keep_next=True)
    style_para(d, "TOC Heading", size=16, bold=True, align=C, before=24, after=24, page_break=True)

    style_para(d, "Title", size=20, bold=True, align=C)
    style_para(d, "Subtitle", size=13, italic=True, align=C)
    for cap in ("Caption", "Image Caption", "Table Caption"):
        style_para(d, cap, size=11, italic=False, align=C, before=4, after=12, spacing=1.0)
    # Table captions sit above the table: keep them with it.
    d.styles["Table Caption"].paragraph_format.keep_with_next = True
    for fig in ("Figure", "Captioned Figure"):
        style_para(d, fig, align=C, after=0, spacing=1.0, keep_next=True)

    d.styles["Hyperlink"].font.color.rgb = BLACK
    d.styles["Hyperlink"].font.underline = False

    sec = d.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.left_margin = Inches(1.5)
    sec.right_margin = sec.top_margin = sec.bottom_margin = Inches(1)

    d.save(OUT)
    print("wrote", OUT)


if __name__ == "__main__":
    main()
