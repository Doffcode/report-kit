"""Post-process the Pandoc/Quarto DOCX into a college-format report.

Pandoc can't do these, so they happen here:
  - cover and front pages (certificate, declaration, ...) from the format profile
  - three sections: cover (no number), front matter (i, ii, ...), body (1, 2, ...)
  - TOC, List of Figures, List of Tables as real Word fields after the abstract
  - "CHAPTER 1 / INTRODUCTION" chapter headings, bold caption labels
The field contents are filled by lo_refresh.py afterwards.
"""
import copy
import re
from pathlib import Path

import sys
import warnings

import docx
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, str(Path(__file__).parent))
warnings.filterwarnings("ignore", message="style lookup by style_id")
CHAPTER_RE = re.compile(r"^\s*(\d+)\.?\s+(.*)$")
CAPTION_STYLES = {"Image Caption", "Table Caption", "Caption"}


# ── small XML helpers ────────────────────────────────────────
def add_field(paragraph, instr, placeholder=""):
    """Append a complex field (PAGE, TOC ...) to a paragraph."""
    def fld(kind):
        r = OxmlElement("w:r")
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), kind)
        if kind == "begin":
            f.set(qn("w:dirty"), "true")
        r.append(f)
        return r
    paragraph._p.append(fld("begin"))
    r = OxmlElement("w:r")
    t = OxmlElement("w:instrText")
    t.set(qn("xml:space"), "preserve")
    t.text = f" {instr} "
    r.append(t)
    paragraph._p.append(r)
    paragraph._p.append(fld("separate"))
    if placeholder:
        paragraph.add_run(placeholder)
    paragraph._p.append(fld("end"))


def page_break_before(paragraph):
    paragraph.paragraph_format.page_break_before = True


def set_section(sect_pr, fmt=None, start=None, title_pg=False):
    """Configure a w:sectPr: page number format/restart and type nextPage."""
    for old in sect_pr.findall(qn("w:pgNumType")) + sect_pr.findall(qn("w:titlePg")):
        sect_pr.remove(old)
    typ = sect_pr.find(qn("w:type"))
    if typ is None:
        typ = OxmlElement("w:type")
        sect_pr.insert(0, typ)
    typ.set(qn("w:val"), "nextPage")
    if fmt or start:
        pg = OxmlElement("w:pgNumType")
        if fmt:
            pg.set(qn("w:fmt"), fmt)
        if start:
            pg.set(qn("w:start"), str(start))
        # schema order: pgNumType comes after pgMar/pgSz; appending before cols is fine for Word/LO
        cols = sect_pr.find(qn("w:cols"))
        if cols is not None:
            cols.addprevious(pg)
        else:
            sect_pr.append(pg)
    if title_pg:
        sect_pr.append(OxmlElement("w:titlePg"))


def end_section_at(paragraph, body_sect_pr):
    """Make `paragraph` the last paragraph of a section (copying page setup)."""
    sp = copy.deepcopy(body_sect_pr)
    for ref in sp.findall(qn("w:headerReference")) + sp.findall(qn("w:footerReference")):
        sp.remove(ref)
    ppr = paragraph._p.get_or_add_pPr()
    ppr.append(sp)
    return sp


# ── front pages (from the format profile) ────────────────────
ALIGN = {"center": WD_ALIGN_PARAGRAPH.CENTER, "left": WD_ALIGN_PARAGRAPH.LEFT,
         "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
MARKUP = re.compile(r"(\*[^*]+\*|_[^_]+_|\[\[TODO:[^\]]*\]\])")


def length_pt(v):
    """'0.8cm' / '1.5in' / '12pt' / '3mm' -> points."""
    m = re.fullmatch(r"(-?[\d.]+)(pt|cm|mm|in)", str(v).strip())
    if not m:
        return 0.0
    n, unit = float(m.group(1)), m.group(2)
    return n * {"pt": 1, "cm": 28.3465, "mm": 2.83465, "in": 72}[unit]


def add_markup(p, text, size=None, bold=False, italic=False):
    """Profile text uses *bold*, _italic_ and [[TODO: ...]] (highlighted)."""
    for part in MARKUP.split(text):
        if not part:
            continue
        b, i, todo = bold, italic, False
        if part.startswith("*") and part.endswith("*") and len(part) > 1:
            part, b = part[1:-1], True
        elif part.startswith("_") and part.endswith("_") and len(part) > 1:
            part, i = part[1:-1], True
        elif part.startswith("[[TODO"):
            part, todo = part[2:-2], True      # show "TODO: ..." like the PDF does
        r = p.add_run(part)
        r.bold, r.italic = b, i
        if size:
            r.font.size = Pt(size)
        if todo:
            highlight(r)


def highlight(run):
    rpr = run._r.get_or_add_rPr()
    h = OxmlElement("w:highlight")
    h.set(qn("w:val"), "yellow")
    rpr.append(h)


class Builder:
    """Creates paragraphs and inserts them before an anchor element."""

    def __init__(self, doc, anchor):
        self.doc, self.anchor = doc, anchor

    def para(self, text="", size=12, bold=False, italic=False, align="center",
             before=0, after=6, style="Normal", spacing=None):
        p = self.doc.add_paragraph(style=style)
        p.alignment = ALIGN[align]
        pf = p.paragraph_format
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        if spacing:
            pf.line_spacing = spacing
        if text:
            add_markup(p, text, size, bold, italic)
        self.anchor.addprevious(p._p)
        return p


TEXT_HEIGHT_PT = 697   # A4 minus 1in top/bottom margins; refined from the profile in build_front_pages


def build_front_pages(doc, prof, anchor, report_dir):
    """Cover (profile.cover) and front pages (profile.pages). Returns the last
    cover paragraph, where the cover section ends."""
    b = Builder(doc, anchor)
    page = prof.get("page") or {}
    margin = page.get("margin") or {}
    text_height = 841.9 - length_pt(margin.get("top", "1in")) - length_pt(margin.get("bottom", "1in"))

    # Cover: Word has no "fill remaining space", so estimate the height of every
    # line and share what's left among the {fill} items.
    items = prof.get("cover") or []
    used, fills = 0.0, 0
    for it in items:
        if "text" in it:
            used += it.get("size", 12) * 1.25 + 4
        elif "space" in it:
            used += length_pt(it["space"])
        elif "logo" in it:
            used += length_pt(it["logo"]) * 1.05 + 6
        elif it.get("fill"):
            fills += 1
    fill_pt = max(0.0, (text_height - used - 40) / fills) if fills else 0
    pending = 0.0
    last = None
    for it in items:
        if "space" in it:
            pending += length_pt(it["space"])
        elif it.get("fill"):
            pending += fill_pt
        elif "text" in it:
            text = it["text"].upper() if it.get("upper") else it["text"]
            last = b.para(text, size=it.get("size", 12), bold=it.get("bold", False),
                          italic=it.get("italic", False), before=pending, after=4, spacing=1.0)
            pending = 0.0
        elif "logo" in it:
            logo = report_dir / it["path"]
            if logo.exists():
                last = b.para(before=pending, after=4, spacing=1.0)
                last.add_run().add_picture(str(logo), width=Pt(length_pt(it["logo"])))
            pending = 0.0
    cover_end = last or b.para()

    def signatures(people):
        if not people:
            return
        t = doc.add_table(rows=1, cols=max(len(people), 2) if len(people) == 1 else len(people))
        cells = t.rows[0].cells
        targets = [cells[-1]] if len(people) == 1 else cells      # a lone signature sits on the right
        for cell, person in zip(targets, people):
            cp = cell.paragraphs[0]
            cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cp.paragraph_format.space_before = Pt(60)
            cp.add_run("______________________").add_break()
            add_markup(cp, person["name"])
            cp.runs[-1].add_break()
            add_markup(cp, person["role"], size=11)
        anchor.addprevious(t._tbl)

    for pg in prof.get("pages") or []:
        t = b.para(pg["title"].upper(), size=prof.get("headings", {}).get("h1", 16), bold=True, after=24)
        page_break_before(t)
        if pg.get("image") and (report_dir / pg["image"]).exists():
            ip = b.para(after=6)
            ip.add_run().add_picture(str(report_dir / pg["image"]), height=Pt(text_height * 0.75))
        if pg.get("text"):
            b.para(pg["text"], align="justify")
        signatures(pg.get("sign") or [])
        if pg.get("place") is not None:
            p = b.para(align="left", before=24)
            add_markup(p, f"Place: {pg['place']}\t\t\t\t\t\tDate:")
    return cover_end


LIST_MARKERS = {"Image Caption": "[[KIT-LIST:Image Caption]]", "Table Caption": "[[KIT-LIST:Table Caption]]"}


def build_lists(doc, anchor, has_figs, has_tables, order=("toc", "lof", "lot"), h1=16):
    """TOC as a Word field; LOF/LOT as markers that lo_refresh.py turns into
    caption-style indexes (LibreOffice drops TOC \\t fields on import)."""
    b = Builder(doc, anchor)
    for item in order:
        if item == "toc":
            b.para("TABLE OF CONTENTS", style="TOC Heading", size=h1, bold=True)
            add_field(b.para(align="left"), r'TOC \o "1-3" \h \z \u', "Right-click and choose Update Field.")
        elif item == "lof" and has_figs:
            b.para("LIST OF FIGURES", style="TOC Heading", size=h1, bold=True)
            b.para(LIST_MARKERS["Image Caption"], align="left")
        elif item == "lot" and has_tables:
            b.para("LIST OF TABLES", style="TOC Heading", size=h1, bold=True)
            b.para(LIST_MARKERS["Table Caption"], align="left")
    return b.para()  # empty paragraph that ends the front-matter section


def abbr_end_paragraph(doc, abbr, first_chapter):
    """Empty paragraph just before chapter 1, closing the front-matter section."""
    p = doc.add_paragraph()
    first_chapter._p.addprevious(p._p)
    return p


def keep_floats_together(doc):
    """Word lets a figure's layout table break between rows, which strands a
    caption on the next page. Forbid row splits and chain every paragraph in
    a float to the next one."""
    for tbl in doc.element.body.iter(qn("w:tbl")):
        paras = [p for p in all_paragraphs(doc) if tbl in p._p.iterancestors()]
        if not any(p.style.name in CAPTION_STYLES for p in paras):
            continue
        for tr in tbl.iter(qn("w:tr")):
            tr_pr = tr.find(qn("w:trPr"))
            if tr_pr is None:
                tr_pr = OxmlElement("w:trPr")
                tr.insert(0, tr_pr)
            if tr_pr.find(qn("w:cantSplit")) is None:
                tr_pr.append(OxmlElement("w:cantSplit"))
        for p in paras[:-1]:
            p.paragraph_format.keep_with_next = True


def fit_images(doc, body_sect_pr):
    """Scale down any picture wider than the text block (minus table-cell
    padding); Word and LibreOffice clip, rather than shrink, oversize images."""
    pg = body_sect_pr.find(qn("w:pgSz"))
    mar = body_sect_pr.find(qn("w:pgMar"))
    twips = int(pg.get(qn("w:w"))) - int(mar.get(qn("w:left"))) - int(mar.get(qn("w:right")))
    max_cx = int(twips * 635 * 0.95)        # 1 twip = 635 EMU; 5% for cell padding
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    for drawing in doc.element.body.iter(qn("w:drawing")):
        extent = drawing.find(f".//{WP}extent")
        if extent is None or int(extent.get("cx")) <= max_cx:
            continue
        k = max_cx / int(extent.get("cx"))
        for el in [extent] + list(drawing.iter(f"{A}ext")):
            el.set("cx", str(int(int(el.get("cx")) * k)))
            el.set("cy", str(int(int(el.get("cy")) * k)))


def border_data_tables(doc):
    """Quarto wraps every figure/table float in a 1-cell layout table. Only
    real data tables (no nested table, no caption inside) get borders."""
    from make_reference_docx import table_borders
    for tbl in doc.element.body.iter(qn("w:tbl")):
        if tbl.find(".//" + qn("w:tbl")) is not None:
            continue
        styles = {s.get(qn("w:val")) for s in tbl.iter(qn("w:pStyle"))}
        if styles & {"ImageCaption", "TableCaption", "Caption", "CaptionedFigure", "Figure"}:
            continue
        if tbl.find(".//" + qn("w:drawing")) is not None:
            continue
        tbl_pr = tbl.find(qn("w:tblPr"))
        table_borders(tbl_pr)
        jc = tbl_pr.find(qn("w:jc"))
        if jc is None:
            jc = OxmlElement("w:jc")
            tbl_pr.append(jc)
        jc.set(qn("w:val"), "center")


# ── XML repair ───────────────────────────────────────────────
# Child order Word requires inside <w:pPr> (ECMA-376 CT_PPr sequence).
PPR_ORDER = ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
             "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
             "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
             "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
             "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
             "rPr", "sectPr", "pPrChange"]
PPR_RANK = {qn(f"w:{t}"): i for i, t in enumerate(PPR_ORDER)}


def repair_ppr(doc):
    """Quarto emits captions with two <w:pPr> blocks, the style in the second
    one and out of order. Word may refuse or 'repair' such a file, and
    python-docx/LibreOffice don't see the style. Merge and re-order."""
    for p in doc.element.body.iter(qn("w:p")):
        pprs = p.findall(qn("w:pPr"))
        if not pprs:
            continue
        base = pprs[0]
        for extra in pprs[1:]:
            for child in list(extra):
                old = base.find(child.tag)
                if old is not None:
                    base.remove(old)
                base.append(child)
            p.remove(extra)
        kids = sorted(base, key=lambda c: PPR_RANK.get(c.tag, len(PPR_ORDER)))
        for c in kids:
            base.remove(c)
        for c in kids:
            base.append(c)
        if p.index(base) != 0:
            p.remove(base)
            p.insert(0, base)


def classify_captions(doc):
    """Quarto styles every float caption 'Image Caption'. Give tables and
    algorithms their own styles so the List of Figures/Tables stay separate,
    and drop Quarto's local alignment/spacing so the style governs."""
    for p in all_paragraphs(doc):
        if p.style.name != "Image Caption":
            continue
        t = p.text
        if t.startswith("Table"):
            p.style = doc.styles["Table Caption"]
        elif t.startswith("Algorithm"):
            p.style = doc.styles["Caption"]
        ppr = p._p.pPr
        for tag in ("w:jc", "w:spacing"):
            el = ppr.find(qn(tag))
            if el is not None:
                ppr.remove(el)


# ── body fixes ───────────────────────────────────────────────
def all_paragraphs(doc):
    """Every paragraph, including those inside Quarto's float layout tables."""
    from docx.text.paragraph import Paragraph
    return [Paragraph(p, doc.part) for p in doc.element.body.iter(qn("w:p"))]


def left_align_algorithms(doc):
    """Algorithm steps sit in a centred float cell; steps must read left-aligned."""
    for tbl in doc.element.body.iter(qn("w:tbl")):
        paras = [p for p in all_paragraphs(doc) if tbl in p._p.iterancestors()]
        if not any(p.style.name in CAPTION_STYLES and p.text.startswith("Algorithm") for p in paras):
            continue
        for p in paras:
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Pt(0)
def style_chapter_headings(doc, chapter):
    """Chapter headings as the profile asks: block ("CHAPTER 1" over the title,
    centred) or inline ("1  INTRODUCTION", left)."""
    first_chapter = None
    style = chapter.get("style", "block")
    upper = chapter.get("upper", True)
    for p in doc.paragraphs:
        if p.style.name != "Heading 1":
            continue
        text = p.text.strip()
        runs = [r for r in p._p if r.tag == qn("w:r") or r.tag == qn("w:hyperlink")]
        for r in runs:
            p._p.remove(r)
        m = CHAPTER_RE.match(text)
        if m:
            if first_chapter is None:
                first_chapter = p
            title = m.group(2).upper() if upper else m.group(2)
            if style == "inline":
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.add_run(f"{m.group(1)}\u2003{title}")
            else:
                r = p.add_run(f"{chapter.get('label', 'CHAPTER')} {m.group(1)}")
                r.add_break(WD_BREAK.LINE)
                p.add_run(title)
            p.paragraph_format.page_break_before = bool(chapter.get("new_page", True))
        else:
            p.add_run(text.upper())
            p.paragraph_format.page_break_before = True
    return first_chapter


def apply_profile_styles(doc, prof):
    """Font, size, line spacing and heading sizes from the profile, so a thesis
    profile with 2.0 spacing or 11pt text changes the DOCX like the PDF."""
    from make_reference_docx import set_font
    family = (prof.get("font") or {}).get("family", "Times New Roman")
    size = (prof.get("font") or {}).get("size", 12)
    spacing = float(prof.get("spacing", 1.5))
    after = prof.get("paragraph_after", 6)
    names = {s.name for s in doc.styles}
    for name in ("Normal", "Body Text", "First Paragraph", "Compact", "Abstract", "Bibliography", "Block Text"):
        if name in names:
            st = doc.styles[name]
            set_font(st.element, family)
            st.font.size = Pt(size)
            st.paragraph_format.line_spacing = spacing
            if name != "Compact":
                st.paragraph_format.space_after = Pt(after)
    heads = prof.get("headings") or {}
    for level, key in ((1, "h1"), (2, "h2"), (3, "h3")):
        name = f"Heading {level}"
        if name in names:
            set_font(doc.styles[name].element, family)
            doc.styles[name].font.size = Pt(heads.get(key, {1: 16, 2: 14, 3: 12}[level]))
    for name in ("Image Caption", "Table Caption", "Caption", "TOC Heading", "Title", "Subtitle"):
        if name in names:
            set_font(doc.styles[name].element, family)
    defaults = doc.styles.element.find(qn("w:docDefaults"))
    fonts = defaults.find(".//" + qn("w:rFonts")) if defaults is not None else None
    if fonts is not None:
        for slot in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            fonts.set(qn(slot), family)


def apply_page_setup(doc, prof):
    """Paper size and margins from the profile on every section."""
    from docx.shared import Mm
    page = prof.get("page") or {}
    margin = page.get("margin") or {}
    w, h = {"a4": (210, 297), "us-letter": (215.9, 279.4), "letter": (215.9, 279.4)}.get(
        str(page.get("paper", "a4")).lower(), (210, 297))
    for sec in doc.sections:
        sec.page_width, sec.page_height = Mm(w), Mm(h)
        for side in ("left", "right", "top", "bottom"):
            if side in margin:
                setattr(sec, f"{side}_margin", Pt(length_pt(margin[side])))


def highlight_todos(doc):
    """Runs in the 'Todo' character style (from filters/todo.lua) get a yellow highlight."""
    todo_ids = {s.style_id for s in doc.styles if s.name == "Todo"}
    for r in doc.element.body.iter(qn("w:r")):
        rs = r.find(qn("w:rPr") + "/" + qn("w:rStyle"))
        if rs is not None and rs.get(qn("w:val")) in todo_ids | {"Todo"}:
            rpr = r.find(qn("w:rPr"))
            h = OxmlElement("w:highlight")
            h.set(qn("w:val"), "yellow")
            rpr.append(h)


def bold_caption_labels(doc):
    """Bold only the "Figure 2.1:" label, splitting the run at the colon."""
    for p in all_paragraphs(doc):
        if p.style.name not in CAPTION_STYLES:
            continue
        for r in p.runs:
            if ":" not in r.text:
                r.bold = True
                continue
            head, tail = r.text.split(":", 1)
            r.text = head + ":"
            r.bold = True
            rest = copy.deepcopy(r._r)
            r._r.addnext(rest)
            from docx.text.run import Run
            rest_run = Run(rest, p)
            rest_run.text = tail
            rest_run.bold = False
            break


def add_page_footer(section, align="center"):
    section.footer.is_linked_to_previous = False
    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align == "right" else WD_ALIGN_PARAGRAPH.CENTER
    add_field(fp, "PAGE", "1")


def drop_title_block(doc):
    """Quarto emits Title/Subtitle/Date paragraphs; the cover replaces them."""
    for p in list(doc.paragraphs):
        if p.style.name in ("Title", "Subtitle", "Date", "Author"):
            p._p.getparent().remove(p._p)


NUMFMT = {"roman": "lowerRoman", "arabic": "decimal"}


def update_fields_on_open(docx_path):
    """Without LibreOffice to pre-fill the TOC, have Word refresh all fields on open."""
    doc = docx.Document(docx_path)
    settings = doc.settings.element
    el = settings.find(qn("w:updateFields"))
    if el is None:
        el = OxmlElement("w:updateFields")
        settings.append(el)
    el.set(qn("w:val"), "true")
    doc.save(docx_path)


def fix(docx_path, prof, report_dir):
    """prof: the format profile with all text already filled for DOCX (profile.resolve)."""
    doc = docx.Document(docx_path)
    repair_ppr(doc)
    classify_captions(doc)
    drop_title_block(doc)
    apply_profile_styles(doc, prof)
    body = doc.element.body
    first_block = body[0]
    body_sect_pr = body.find(qn("w:sectPr"))

    paras = all_paragraphs(doc)
    has_figs = any(p.style.name == "Image Caption" for p in paras)
    has_tables = any(p.style.name == "Table Caption" for p in paras)

    first_chapter = style_chapter_headings(doc, prof.get("chapter") or {})
    bold_caption_labels(doc)
    border_data_tables(doc)
    left_align_algorithms(doc)
    keep_floats_together(doc)
    highlight_todos(doc)
    apply_page_setup(doc, prof)
    fit_images(doc, body_sect_pr)
    cover_end = build_front_pages(doc, prof, first_block, report_dir)
    abbr = next((p for p in doc.paragraphs if p.style.name == "Heading 1"
                 and p.text.strip() == "LIST OF ABBREVIATIONS"), None)
    lists_at = abbr or first_chapter
    front_end = build_lists(doc, lists_at._p if lists_at is not None else body_sect_pr,
                            has_figs, has_tables, prof.get("front_lists") or ("toc", "lof", "lot"),
                            (prof.get("headings") or {}).get("h1", 16))
    if abbr is not None:   # lists now come first; abbreviations follow on a new page
        abbr.paragraph_format.page_break_before = True
        front_end = abbr_end_paragraph(doc, abbr, first_chapter)

    # Sections: cover | front matter (roman from i) | body (arabic from 1)
    nums = (prof.get("page") or {}).get("numbers") or {}
    set_section(end_section_at(cover_end, body_sect_pr))
    set_section(end_section_at(front_end, body_sect_pr), fmt=NUMFMT.get(nums.get("front"), "lowerRoman"), start=1)
    set_section(body_sect_pr, fmt=NUMFMT.get(nums.get("body"), "decimal"), start=1)
    # A section break already starts a new page; drop the duplicate break on chapter 1.
    if first_chapter is not None:
        first_chapter.paragraph_format.page_break_before = False

    secs = doc.sections
    secs[0].footer.is_linked_to_previous = False   # cover: empty footer
    add_page_footer(secs[1], nums.get("align", "center"))
    secs[2].footer.is_linked_to_previous = True    # same PAGE field, format from pgNumType
    doc.save(docx_path)


if __name__ == "__main__":
    raise SystemExit("run via `report build`")
