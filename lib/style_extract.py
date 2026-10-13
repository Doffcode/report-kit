"""Draft a format profile from a sample report PDF.

    report style-extract sample.pdf --name mycollege [--force]

Measures paper, margins, body font/size, line spacing, justification,
heading sizes, chapter heading style, caption labels, page numbering, the
order of front pages and lists, certificate/declaration wording and the
cover layout, then writes profiles/<name>.yml (extends: project). Values it
could not measure are left to the project profile. Always build a report
with the new profile and compare it with the sample before trusting it.
Scanned PDFs (no text layer) can't be measured.
"""
import re
import statistics
from collections import Counter
from pathlib import Path

import pymupdf
import yaml

KIT = Path(__file__).resolve().parent.parent

FRONT_TITLES = {
    "certificate": r"(bonafide\s+)?certificate", "declaration": r"declaration|candidate'?s declaration",
    "acknowledgement": r"acknowledge?ments?", "abstract": r"abstract|synopsis",
    "toc": r"(table of )?contents", "lof": r"list of figures", "lot": r"list of tables",
    "abbreviations": r"list of (abbreviations|acronyms|symbols( and abbreviations)?)|abbreviations|nomenclature",
}
FONT_NAMES = [("times", "Times New Roman"), ("arial", "Arial"), ("helvetica", "Arial"), ("calibri", "Calibri"),
              ("cambria", "Cambria"), ("georgia", "Georgia"), ("garamond", "Garamond"),
              ("bookantiqua", "Book Antiqua"), ("liberationserif", "Times New Roman"),
              ("liberationsans", "Arial"), ("computermodern", "New Computer Modern"), ("cmr", "New Computer Modern")]


def family(fontname):
    key = re.sub(r"^[A-Z]{6}\+", "", fontname).lower().replace(" ", "").replace("-", "")
    return next((nice for k, nice in FONT_NAMES if k in key), re.sub(r"[-,].*$", "", fontname.split("+")[-1]))


def is_bold(span):
    return bool(span["flags"] & 16) or "bold" in span["font"].lower()


def lines_of(page):
    """[(text, size, bold, fontname, bbox)] per visual line. Pieces on the same
    baseline (e.g. an inline "1" + "INTRODUCTION" heading) are merged."""
    raw = []
    for b in page.get_text("dict")["blocks"]:
        if b["type"] != 0:
            continue
        for ln in b["lines"]:
            spans = [s for s in ln["spans"] if s["text"].strip()]
            if spans:
                raw.append((spans, pymupdf.Rect(ln["bbox"])))
    raw.sort(key=lambda x: (round(x[1].y1), x[1].x0))
    merged = []
    for spans, r in raw:
        if merged and abs(merged[-1][1].y1 - r.y1) < 2 and r.x0 >= merged[-1][1].x1 - 1:
            merged[-1] = (merged[-1][0] + spans, merged[-1][1] | r)
        else:
            merged.append((spans, r))
    out = []
    for spans, r in merged:
        text = " ".join(s["text"].strip() for s in spans)
        main = max(spans, key=lambda s: len(s["text"]))
        out.append(Line(text, round(main["size"], 1), is_bold(main), main["font"], r,
                        bool(main["flags"] & 2) or "italic" in main["font"].lower(), markup(spans)))
    return out


class Line(tuple):
    """(text, size, bold, font, bbox) plus .italic and .marked (text with *bold* runs)."""
    def __new__(cls, text, size, bold, font, bbox, italic, marked):
        t = super().__new__(cls, (text, size, bold, font, bbox))
        t.italic, t.marked = italic, marked
        return t


def markup(spans):
    """Join spans, wrapping bold runs in *...* (the profile's markup)."""
    parts, cur = [], []
    for s in spans:
        txt = s["text"]
        if is_bold(s):
            cur.append(txt)
        else:
            if cur:
                parts.append(f"*{''.join(cur).strip()}* " if "".join(cur).strip() else "".join(cur))
                cur = []
            parts.append(txt)
    if cur and "".join(cur).strip():
        parts.append(f"*{''.join(cur).strip()}*")
    return re.sub(r"\s+", " ", "".join(parts)).strip()


PAGE_NUM = re.compile(r"[-–\s]*([0-9]{1,3}|[ivxlc]{1,6})[-–\s]*")


def nearest(value, options):
    return min(options, key=lambda o: abs(o - value))


def to_len(pt):
    """Points -> a tidy profile length (inches if close to a quarter inch, else cm)."""
    inches = pt / 72
    q = round(inches * 4) / 4
    if abs(inches - q) < 0.04:
        return f"{q:g}in"
    return f"{round(pt / 28.3465, 1):g}cm"


def extract(pdf_path):
    doc = pymupdf.open(pdf_path)
    if sum(len(p.get_text().strip()) for p in doc) < 200:
        raise SystemExit("This PDF has no text layer (scanned?). Style extraction needs a text PDF.")
    W, H = doc[0].rect.width, doc[0].rect.height
    pages = [lines_of(p) for p in doc]
    found, notes = {}, []

    # ── paper ──
    paper = "a4" if abs(W - 595) < 8 and abs(H - 842) < 8 else "us-letter" if abs(W - 612) < 8 else None
    if paper:
        found["page"] = {"paper": paper}

    # ── body font, size: most characters ──
    chars = Counter()
    for pl in pages:
        for text, size, bold, font, _ in pl:
            if not bold:
                chars[(family(font), size)] += len(text)
    (body_font, body_size), _ = chars.most_common(1)[0]
    found["font"] = {"family": body_font, "size": nearest(body_size, [10, 10.5, 11, 11.5, 12, 13, 14])}
    if body_font not in ("Times New Roman", "Arial", "Calibri", "Cambria", "Georgia", "Garamond", "Book Antiqua"):
        notes.append(f"body font “{body_font}” may not be installed; check `typst fonts`")

    # ── where chapters start (to separate front matter from body) ──
    def page_title(pl):
        top = [l for l in pl if l[4].y0 < H * 0.35 and len(l[0]) < 70]
        return " ".join(l[0] for l in top[:3]).lower()

    inline_head = re.compile(r"1\.?\s+[A-Z][A-Za-z ]{3,}")
    chapter_pages = [i for i, pl in enumerate(pages)
                     if re.search(r"\bchapter\s+(1|one|i)\b", page_title(pl))
                     or any(inline_head.fullmatch(l[0]) and l[1] > body_size + 1 for l in pl[:4])]
    first_body = chapter_pages[0] if chapter_pages else max(1, len(doc) // 5)
    body_pages = [pl for pl in pages[first_body:] if pl]

    # ── margins: left/right from the most common line starts/ends (justified text
    #    ends at one x); top/bottom from each page's first and last line ──
    starts, ends, tops, bottoms = Counter(), Counter(), [], []
    long_ends = []
    for pl in body_pages:
        content = [l for l in pl if not PAGE_NUM.fullmatch(l[0])]
        for l in content:
            if abs(l[1] - body_size) < 0.6 and len(l[0]) > 40:
                starts[round(l[4].x0)] += 1
                ends[round(l[4].x1)] += 1
                long_ends.append(l[4].x1)
        if content:
            tops.append(min(l[4].y0 for l in content))
            bottoms.append(max(l[4].y1 for l in content))
    if starts and tops:
        left = starts.most_common(1)[0][0]
        right_edge = ends.most_common(1)[0][0]
        top = sorted(tops)[len(tops) // 10]                       # pages that start at the margin
        bottom = sorted(bottoms)[int(len(bottoms) * 0.9) - 1]     # full pages end at the margin
        found.setdefault("page", {})["margin"] = {
            "left": to_len(left), "right": to_len(W - right_edge),
            "top": to_len(top), "bottom": to_len(H - bottom)}
        # justified: most long lines end exactly at the common right edge
        near_full = [x for x in long_ends if x > right_edge - 40]
        found["justify"] = sum(1 for x in near_full if abs(x - right_edge) < 4) / max(len(near_full), 1) > 0.6

    # ── line spacing: baseline distance of consecutive body lines ──
    gaps = []
    for pl in body_pages:
        main = [l for l in pl if abs(l[1] - body_size) < 0.6]
        for a, b in zip(main, main[1:]):
            d = b[4].y1 - a[4].y1
            if 0.9 * body_size < d < 3 * body_size and abs(a[4].x0 - b[4].x0) < 40:
                gaps.append(d)
    if gaps:
        ratio = statistics.median(gaps) / (body_size * 1.149)
        found["spacing"] = nearest(ratio, [1.0, 1.15, 1.5, 2.0])

    # ── headings: bold sizes above body, most common first ──
    head_sizes = Counter(l[1] for pl in body_pages for l in pl if l[2] and l[1] > body_size + 0.4 and len(l[0]) < 90)
    sizes = sorted({s for s, n in head_sizes.items() if n >= 2}, reverse=True)
    if sizes:
        hs = {"h1": sizes[0], "h2": sizes[1] if len(sizes) > 1 else body_size + 2, "h3": body_size}
        found["headings"] = {k: nearest(v, [12, 13, 14, 15, 16, 17, 18, 20, 22, 24]) for k, v in hs.items()}

    # ── chapter style ──
    label = None
    for pl in pages[first_body:]:
        for l in pl[:4]:
            m = re.fullmatch(r"(CHAPTER|Chapter)\s*[-–:]?\s*(\d+|[IVX]+|One|ONE)", l[0].strip())
            if m:
                label = m.group(1)
                break
        if label:
            break
    if label:
        found["chapter"] = {"style": "block", "label": label, "new_page": True}
    elif chapter_pages:
        first = next(l for l in pages[chapter_pages[0]][:4] if inline_head.fullmatch(l[0]))
        found["chapter"] = {"style": "inline", "new_page": True,
                            "upper": first[0].split(None, 1)[1].isupper()}

    # ── captions ──
    cap = Counter()
    for pl in pages:
        for l in pl:
            m = re.match(r"^(Fig\.|Figure|FIGURE|Fig)\s*\d+([.\-]\d+)?\s*[:.]", l[0])
            if m:
                cap[m.group(1)] += 1
    if cap:
        found["captions"] = {"figure": cap.most_common(1)[0][0], "table": "Table"}

    # ── page numbers ──
    nums = []
    for i, pl in enumerate(pages):
        for text, size, bold, font, r in pl:
            if (r.y0 > H * 0.85 or r.y1 < H * 0.08) and PAGE_NUM.fullmatch(text.strip()):
                nums.append((i, text.strip(" -–"), (r.x0 + r.x1) / 2))
    if nums:
        roman = [n for n in nums if re.fullmatch(r"[ivxlc]+", n[1])]
        align = "right" if statistics.median(n[2] for n in nums) > W * 0.7 else "center"
        found["page"] = {**found.get("page", {}), "numbers": {
            "front": "roman" if roman else "arabic", "body": "arabic", "align": align}}

    # ── front matter: which pages, in which order ──
    order, front_texts = [], {}
    for i, pl in enumerate(pages[1:first_body], 1):
        heads = [l for l in pl if l[4].y0 < H * 0.3 and (l[2] or l[1] > body_size + 1) and len(l[0]) < 60]
        for h in heads[:2]:
            for key, pat in FRONT_TITLES.items():
                if re.fullmatch(pat, h[0].strip().lower().rstrip(":")):
                    if key not in order:
                        order.append(key)
                        front_texts[key] = (h[0].strip(), i)
                    break
    lists = [k for k in order if k in ("toc", "lof", "lot")]
    if lists:
        found["front_lists"] = lists
    found["abbreviations"] = "abbreviations" in order

    # ── cover: every line on page 1, with spacing ──
    cover_lines = sorted(pages[0], key=lambda l: l[4].y0)
    title_line = max(cover_lines, key=lambda l: l[1]) if cover_lines else None
    cover, prev_y = [], None
    imgs = [b for b in doc[0].get_text("dict")["blocks"] if b["type"] == 1]
    events = [("line", l[4].y0, l) for l in cover_lines] + [("img", b["bbox"][1], b) for b in imgs]
    title_text = None
    for kind, y, item in sorted(events, key=lambda e: e[1]):
        if prev_y is not None:
            gap = y - prev_y
            if gap > H * 0.12:
                cover.append({"fill": True})
            elif gap > body_size * 1.6:
                cover.append({"space": f"{round((gap - body_size) / 28.35, 1):g}cm"})
        if kind == "img":
            w = item["bbox"][2] - item["bbox"][0]
            cover.append({"logo": f"{round(w / 28.35, 1):g}cm"})
            prev_y = item["bbox"][3]
            continue
        text, size, bold, font, r = item
        entry = {"text": text, "size": nearest(size, [10, 11, 12, 13, 14, 15, 16, 18, 20, 22, 24, 26, 28])}
        if bold:
            entry["bold"] = True
        if item.italic:
            entry["italic"] = True
        if text.isupper() and len(text) > 3:
            entry["upper"] = True
        if item is title_line:
            title_text = text
        cover.append(entry)
        prev_y = r.y1
    found["cover"], literals = placeholderise_cover(cover, title_text)
    rtype = literals.pop("__report_type__", None)
    if rtype:
        found["report_type"] = rtype
        kind = re.sub(r"(?i)^an?\s+", "", rtype).lower()
        found["doc_kind"] = kind

    # ── certificate / declaration wording: the paragraph(s) above the big gap
    #    that precedes the signatures ──
    pages_out = []
    for key in ("certificate", "declaration"):
        if key not in front_texts:
            continue
        heading, idx = front_texts[key]
        body = sorted([l for l in pages[idx] if abs(l[1] - body_size) < 0.8 and l[4].y0 > H * 0.1
                       and l[0].strip() != heading], key=lambda l: l[4].y0)
        para, prev, pitch = [], None, None
        for l in body:
            if prev is not None:
                gap = l[4].y0 - prev
                pitch = pitch or gap
                if gap > 2.2 * pitch:
                    break
            para.append(l.marked)
            prev = l[4].y0
        text = " ".join(para)
        text = re.sub(r"\s+([,.;:])", r"\1", text).replace("* *", " ")
        entry = {"title": heading.title(), "text": placeholderise_text(text, title_text, literals)}
        entry["sign"] = ["guide", "hod"] if key == "certificate" else ["student"]
        entry["place_date"] = any(re.search(r"(?i)\b(place|date)\b", l[0]) for l in pages[idx])
        pages_out.append(entry)
    if pages_out:
        found["pages"] = pages_out
    return found, notes, {"pages": len(doc), "first_chapter_page": first_body + 1, "front_order": order}


# ── turning the sample's literal text into {placeholders} ─────
COVER_RULES = [
    (r"(?i)^(an?\s+)?((b\.?\s?tech|m\.?\s?tech|mini|major|final[- ]year|summer|industrial)\s+)*(project report|thesis|dissertation|internship report|seminar report|training report|report)$", "{report_type}"),
    (r"(?i)^(bachelor|master|doctor)\b.*", "*{degree}*"),
    (r"(?i).*\b(department|school) of\b.*", "*{department}*"),
    (r"(?i).*\b(institute|university|college)\b.*", "*{college}*"),
    (r"(?i)^(january|february|march|april|may|june|july|august|september|october|november|december)\s+\d{4}$", "{date}"),
    (r"(?i)^\(?\s*(roll|reg(istration)?|enrol?l?ment)\.?\s*(no|number)?\.?\s*[:\-]?\s*[A-Z0-9/]+\s*\)?$", "ROLL"),
    (r"^\(?[0-9]{6,}[A-Z0-9]*\)?$|^\(?[A-Z]{2,}[0-9]{5,}\)?$", "ROLL"),
]
AFTER = {r"(?i)^(submitted\s+)?by:?$": "*{name}*", r"(?i).*(guidance|supervision) of:?$": "*{guide}*",
         r"(?i)^in$": "*{branch}*"}


def placeholderise_cover(cover, title_text):
    """Replace the sample's own details with {placeholders}. Returns the new cover
    and a map of literal text -> placeholder, reused for the certificate wording."""
    out, expect, literals = [], None, {}
    for item in cover:
        if "text" not in item:
            out.append(item)
            continue
        t = item["text"].strip()
        ph = None
        if t == title_text:
            ph = "{title}"
        elif expect:
            ph = expect
            expect = "{guide_title}" if expect == "*{guide}*" else None
        else:
            for pat, cand in COVER_RULES:
                if re.fullmatch(pat, t):
                    ph = cand
                    break
        if ph == "ROLL":       # keep "Roll No." etc.; swap only the number itself
            ident = re.findall(r"[A-Z0-9/]*\d{4,}[A-Z0-9/]*", t)[-1]
            literals[ident] = "{roll}"
            item = {**item, "text": t.replace(ident, "{roll}")}
        elif ph:
            item = {**item, "text": ph}
            if ph == "{report_type}":
                literals["__report_type__"] = t
                item.pop("bold", None) if not item.get("bold") else None
            elif ph not in ("{title}", "{date}"):
                literals[re.sub(r"^\(|\)$", "", t)] = ph.strip("*").strip("()")
        for pat, nxt in AFTER.items():
            if re.fullmatch(pat, t):
                expect = nxt
        out.append(item)
    return out, literals


def placeholderise_text(text, title_text, literals):
    if title_text:
        text = re.sub(re.escape(title_text), "{title}", text, flags=re.I)
    for lit, ph in sorted(literals.items(), key=lambda kv: -len(kv[0])):
        if len(lit) > 3:
            text = text.replace(lit, ph)
    text = re.sub(r"(?i)\b(Bachelor|Master|Doctor) of [A-Z][A-Za-z ]+?(?=[,.]| in\b)", "{degree}", text)
    text = re.sub(r"\b[0-9]{6,}[A-Z0-9]*\b", "{roll}", text)
    return text


def write_profile(name, found, notes, info, force=False):
    dest = KIT / "profiles" / f"{name}.yml"
    if dest.exists() and not force:
        raise SystemExit(f"{dest} exists; use --force to overwrite")
    head = [
        f"# Drafted by `report style-extract` from a sample PDF ({info['pages']} pages, chapter 1 on",
        f"# page {info['first_chapter_page']}; front matter seen: {', '.join(info['front_order']) or 'none'}).",
        "# Values here were MEASURED; anything missing comes from the project profile.",
        "# Check: cover placeholders ({name}, {roll}, ...) and certificate wording, then build a",
        f"# report with `profile: {name}` and compare it with the sample side by side.",
    ] + [f"# NOTE: {n}" for n in notes] + ["extends: project", f"name: {name}"]
    body = yaml.safe_dump(found, sort_keys=False, allow_unicode=True, width=100)
    dest.write_text("\n".join(head) + "\n" + body)
    return dest


def main(pdf, name, force=False):
    found, notes, info = extract(pdf)
    dest = write_profile(name, found, notes, info, force)
    print(f"  wrote {dest}")
    page = found.get("page", {})
    print(f"  paper {page.get('paper', '?')}, margins {page.get('margin', '?')}")
    print(f"  font {found['font']['family']} {found['font']['size']}pt, spacing {found.get('spacing', '?')}, "
          f"justified {found.get('justify', '?')}")
    print(f"  chapters {found.get('chapter', '?')}, headings {found.get('headings', '?')}")
    print(f"  front pages {info['front_order']}, captions {found.get('captions', '?')}, "
          f"page numbers {page.get('numbers', '?')}")
    unfilled = [c["text"] for c in found["cover"] if "text" in c and "{" not in c["text"]]
    if unfilled:
        print(f"  cover lines kept as literal text (check them): {unfilled}")
    for n in notes:
        print(f"  note: {n}")
