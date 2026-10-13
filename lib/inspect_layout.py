"""Layout inspector: the page-by-page look a person would do, automated.

Flags half-empty pages in the middle of a chapter, headings stranded at the
bottom of a page, and anything running past the margins, on the PDF and on
a LibreOffice rendering of the DOCX. Also reports what moved since the last
build and writes a contact sheet (_build/pages.png) for a quick glance.
"""
import json
import re
from pathlib import Path

import pymupdf

# Text-area margins in points; set from the report's format profile in inspect().
MARGIN = {"left": 108, "right": 72, "top": 72, "bottom": 72}
UNDERFULL = 0.55       # a mid-chapter page filled less than this is flagged
TOL = 3                # points of tolerance for margin overflow
MIN_TEXT = 6.0         # pt; smaller text (usually a shrunken diagram) is unreadable in print


CHAPTER_TITLE = re.compile(r"^(\d+\s|CHAPTER\s+\d+)")


def chapter_starts(doc):
    """Physical page indexes where top-level outline entries begin, and the
    first numbered chapter (front matter before it is not inspected for fill)."""
    toc = [(t, p - 1) for lvl, t, p, *_ in doc.get_toc() if lvl == 1 and p > 0]
    first = next((p for t, p in toc if CHAPTER_TITLE.match(t)), None)
    return sorted({p for _, p in toc}), first


def label(doc, i):
    """The printed page number (from the footer), else the physical page."""
    page = doc[i]
    foot = [b[4].strip() for b in page.get_text("blocks") if b[1] > page.rect.height - MARGIN["bottom"]]
    num = next((f for f in foot if re.fullmatch(r"[0-9]+|[ivxlc]+", f)), None)
    return f"p.{num}" if num else f"page {i + 1}"


def content_boxes(page):
    boxes = []
    d = page.get_text("rawdict")
    for b in d["blocks"]:
        if b["type"] == 0:
            for line in b["lines"]:
                for s in line["spans"]:
                    ink = [c for c in s["chars"] if c["c"].strip()]   # trailing spaces don't count
                    if ink:
                        s["text"] = "".join(c["c"] for c in s["chars"])
                        r = pymupdf.Rect(ink[0]["bbox"])
                        for c in ink[1:]:
                            r |= pymupdf.Rect(c["bbox"])
                        boxes.append((r, s))
        else:
            boxes.append((pymupdf.Rect(b["bbox"]), None))
    for dr in page.get_drawings():
        r = dr["rect"]
        if r.width > 1 or r.height > 1:
            boxes.append((r, None))
    return boxes


def inspect_pdf(path: Path, name: str):
    out = []
    if not path.exists():
        return out, {}
    doc = pymupdf.open(path)
    starts, first_body = chapter_starts(doc)
    if first_body is None:
        return out, {}
    last_pages = {s - 1 for s in starts[1:]} | {len(doc) - 1}
    for i in range(len(doc)):
        page = doc[i]
        w, h = page.rect.width, page.rect.height
        area_top, area_bottom = MARGIN["top"], h - MARGIN["bottom"]
        footer = [b for b in page.get_text("blocks") if b[1] > area_bottom + 4 and re.fullmatch(r"\s*([0-9]+|[ivxlc]+)\s*", b[4])]
        boxes = [(r, s) for r, s in content_boxes(page)
                 if not any(pymupdf.Rect(f[:4]).contains(r) for f in footer)]  # skip page number
        if not boxes:
            continue
        for r, s in boxes:
            if r.y0 < area_top - 20 or r.y1 > area_bottom + 4:   # header room above, footer below
                what = f"“{s['text'].strip()[:40]}”" if s else "a figure"
                out.append(("FAIL" if name == "PDF" else "WARN",
                            f"{name} {label(doc, i)}: {what} runs past the top/bottom margin (cut off?)"))
                break
        for r, s in boxes:
            if r.x1 > w - MARGIN["right"] + TOL or r.x0 < MARGIN["left"] - TOL:
                what = f"“{s['text'].strip()[:40]}”" if s else "a figure/table"
                out.append(("FAIL" if name == "PDF" else "WARN",
                            f"{name} {label(doc, i)}: {what} runs past the margin"))
                break
        tiny = [s for _, s in boxes if s and s["size"] < MIN_TEXT and s["text"].strip()]
        if tiny:
            out.append(("WARN", f"{name} {label(doc, i)}: {len(tiny)} text item(s) below {MIN_TEXT:g}pt, "
                                f"e.g. “{tiny[0]['text'].strip()[:30]}” ({tiny[0]['size']:.1f}pt): "
                                "a figure is probably too wide; draw it top-to-bottom"))
        if first_body is not None and i >= first_body:
            ordered = sorted(boxes, key=lambda x: x[0].y0)
            top_r, top_s = ordered[0]
            if top_s and re.match(r"^(Figure|Fig\.)\s*\d+(\.\d+)?:", top_s["text"].strip()):
                out.append(("FAIL" if name == "PDF" else "WARN",
                            f"{name} {label(doc, i)}: caption “{top_s['text'].strip()[:30]}” is separated from its figure"))
            bot_r, bot_s = max(boxes, key=lambda x: x[0].y1)
            if bot_s and re.match(r"^(Table|Algorithm)\s*\d+(\.\d+)?:", bot_s["text"].strip()):
                out.append(("FAIL" if name == "PDF" else "WARN",
                            f"{name} {label(doc, i)}: caption “{bot_s['text'].strip()[:30]}” is separated from its content"))
        if i < first_body or i in last_pages or i + 1 in starts:
            continue
        bottom = max(r.y1 for r, _ in boxes)
        fill = (bottom - area_top) / (area_bottom - area_top)
        if fill < UNDERFULL:
            out.append(("WARN", f"{name} {label(doc, i)}: only {fill:.0%} full in the middle of a chapter"))
        last_spans = [(r, s) for r, s in boxes if s and r.y0 > area_top + 0.85 * (area_bottom - area_top)]
        if last_spans:
            r, s = max(last_spans, key=lambda x: x[0].y1)
            if s["size"] >= 13 and "Bold" in s["font"]:
                out.append(("WARN", f"{name} {label(doc, i)}: heading “{s['text'][:40]}” stranded at page bottom"))
    toc = {t: label(doc, p - 1) for lvl, t, p, *_ in doc.get_toc() if lvl <= 2 and p > 0}
    return out, {"pages": len(doc), "toc": toc}


def contact_sheet(pdf: Path, dest: Path, cols=8, dpi=36):
    from PIL import Image
    doc = pymupdf.open(pdf)
    ims = []
    for p in doc:
        pix = p.get_pixmap(dpi=dpi)
        ims.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
    w, h = ims[0].size
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * w, rows * h), "#888")
    for k, im in enumerate(ims):
        sheet.paste(im, ((k % cols) * w, (k // cols) * h))
    sheet.save(dest)


def inspect(rdir: Path, cfg: dict):
    import profile as fmt
    MARGIN.update(fmt.margins_pt(fmt.load(cfg.get("profile", "project"), rdir, cfg.get("format"))))
    build = rdir / "_build"
    build.mkdir(exist_ok=True)
    out, state = inspect_pdf(rdir / "out/report.pdf", "PDF")
    docx_out, _ = inspect_pdf(build / "docx-preview.pdf", "DOCX")
    out += docx_out

    prev_file = build / "layout.json"
    if state and prev_file.exists():
        prev = json.loads(prev_file.read_text())
        if prev.get("pages") != state["pages"]:
            out.append(("INFO", f"page count {prev.get('pages')} → {state['pages']}"))
        for title, page in state["toc"].items():
            old = prev.get("toc", {}).get(title)
            if old and old != page:
                out.append(("INFO", f"moved: “{title[:50]}” {old} → {page}"))
    if state:
        prev_file.write_text(json.dumps(state, indent=1))
        contact_sheet(rdir / "out/report.pdf", build / "pages.png")
        out.append(("INFO", f"page overview: {(build / 'pages.png').relative_to(rdir.parent)}"))
    return out
