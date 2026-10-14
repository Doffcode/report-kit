"""Slide designs chosen by what a slide says, not only by its format.

A slide's content suggests a design: a sentence and three points (the problem), a
numbered list (objectives), two lists under ### headings (existing vs proposed, or
advantages vs limitations), a diagram with a few notes, a chart with its takeaway, …
`choose()` picks one and says why; a class on the slide ({.compare}) picks one
directly. The designs share one rule: what's drawn says something about the content,
and the accent colour marks the one thing that matters (the winner, the takeaway,
the current section), not decoration.

Icons: `- :lock: **Privacy**: …` (any Lucide icon name, lucide.dev/icons). With
`icons: auto` (default) a group of items gets icons only when every item's label has
an obvious one; otherwise none (half the items with icons looks accidental).
"""
import re
from pathlib import Path

from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Pt

KIT = Path(__file__).resolve().parent.parent
ICON_DIR = KIT / "node/node_modules/lucide-static/icons"
IN = 914400

PATTERNS = ("agenda", "lead", "lead-cards", "numbered", "checklist", "compare", "proscons", "columns",
            "callouts", "insight", "definition", "takeaways", "roadmap")

# label words → icon, only where the picture is obvious
ICON_WORDS = {
    "privacy": "lock", "private": "lock", "secure": "shield-check", "security": "shield-check", "safety": "shield",
    "data": "database", "dataset": "database", "storage": "hard-drive", "server": "server", "servers": "server",
    "client": "smartphone", "clients": "smartphone", "device": "smartphone", "devices": "smartphone",
    "mobile": "smartphone", "phone": "smartphone", "hospital": "hospital", "hospitals": "hospital",
    "cloud": "cloud", "network": "network", "communication": "radio-tower", "bandwidth": "radio-tower",
    "accuracy": "target", "precision": "crosshair", "speed": "zap", "fast": "zap", "faster": "zap",
    "latency": "timer", "time": "clock", "cost": "wallet", "costs": "wallet", "energy": "battery-charging",
    "power": "battery-charging", "model": "brain", "models": "brain", "training": "dumbbell", "learning": "brain",
    "users": "users", "user": "user", "people": "users", "team": "users", "scalability": "trending-up",
    "scalable": "trending-up", "growth": "trending-up", "performance": "gauge", "efficiency": "gauge",
    "search": "search", "survey": "book-open", "literature": "book-open", "research": "microscope",
    "code": "code", "software": "code", "hardware": "cpu", "compute": "cpu", "gpu": "cpu", "memory": "memory-stick",
    "testing": "flask-conical", "experiments": "flask-conical", "experiment": "flask-conical",
    "evaluation": "clipboard-check", "results": "bar-chart-3", "analysis": "line-chart", "visualization": "pie-chart",
    "deployment": "rocket", "launch": "rocket", "web": "globe", "internet": "globe", "api": "plug",
    "database": "database", "camera": "camera", "image": "image", "images": "image", "video": "video",
    "audio": "mic", "speech": "mic", "text": "file-text", "documents": "file-text", "report": "file-text",
    "health": "heart-pulse", "medical": "stethoscope", "agriculture": "sprout", "environment": "leaf",
    "money": "indian-rupee", "finance": "landmark", "education": "graduation-cap", "students": "graduation-cap",
    "location": "map-pin", "maps": "map", "sensor": "radar", "sensors": "radar", "iot": "cpu",
    "averaging": "sigma", "aggregation": "merge", "fairness": "scale", "bias": "scale", "robustness": "shield",
    "drift": "shuffle", "heterogeneity": "shuffle", "noise": "audio-waveform", "privacy-preserving": "lock",
    "automation": "bot", "chatbot": "bot", "ai": "sparkles", "prediction": "telescope", "forecast": "telescope",
}


# ── items and icons ──────────────────────────────────────────
def parse_item(text):
    """'- :lock: **Privacy**: data stays local' → (icon, head, body)."""
    import deck
    icon = None
    m = re.match(r"^:([a-z0-9-]+):\s*(.*)$", text)
    if m:
        icon, text = m.group(1), m.group(2)
    head, body = deck.split_head(text)
    return icon, head, body


def auto_icons(heads, mode):
    """An icon for every label, or for none."""
    if mode in ("off", "manual", "false", False) or not heads or not all(heads):
        return None
    out = []
    for h in heads:
        words = re.findall(r"[a-z][a-z-]*", h.lower())
        hit = next((ICON_WORDS[w] for w in words if w in ICON_WORDS), None) or \
            next((ICON_WORDS[w.rstrip("s")] for w in words if w.rstrip("s") in ICON_WORDS), None)
        if not hit:
            return None
        out.append(hit)
    return out if len(set(out)) == len(out) else None      # two items with the same icon say nothing


def icon_png(b, name, colour):
    """Lucide icon `name` in `colour` → PNG in _build/icons (None if unknown)."""
    src = ICON_DIR / f"{name}.svg"
    if not src.exists():
        msg = (f"icon :{name}: not found (names: lucide.dev/icons)" if ICON_DIR.exists()
               else "icons are not installed: re-run install.sh")
        if ("WARN", msg) not in b.problems:
            b.problems.append(("WARN", msg))
        return None
    out = b.sdir / "_build" / "icons" / f"{name}-{colour}.png"
    if not out.exists():
        import pymupdf
        out.parent.mkdir(parents=True, exist_ok=True)
        svg = src.read_text().replace("currentColor", "#" + colour)
        doc = pymupdf.open(stream=svg.encode(), filetype="svg")
        doc[0].get_pixmap(matrix=pymupdf.Matrix(12, 12), alpha=True).save(out)
    return out


def items_with_icons(b, items):
    """[[text, subs]] → [(icon, head, body, subs)] with auto icons where every label has one."""
    parsed = [parse_item(t) + (subs,) for t, subs in items]
    if not any(p[0] for p in parsed):
        auto = auto_icons([p[1] for p in parsed], str(b.cfg.get("icons", "auto")).lower())
        if auto:
            parsed = [(ic, h, bd, sb) for ic, (_, h, bd, sb) in zip(auto, parsed)]
    return parsed


# ── shared drawing ───────────────────────────────────────────
def title(b, s, sl, n, **kw):
    import deck
    deck.title_text(s.shapes.title, sl["title"], b.tpl, b.cites, b.problems, b.where(n, sl), **kw)


def item_paras(b, head, body, subs):
    import deck
    out = []
    if head:
        out.append((0, [(x, st + "b" if "b" not in st else st) for x, st in deck.runs(head, b.cites)]))
    if body:
        out.append((1 if head else 0, deck.runs(body, b.cites)))
    out += [(1, deck.runs(x, b.cites)) for x in subs]
    return out


def marker(b, s, kind, box, i, icon=None, colour=None):
    """Something at the start of a row: number in a circle, big numeral, icon, ✓ / ✗, dot."""
    import deck
    l, t, w, h = (int(v) for v in box)
    colour = colour or b.ink()["accent"]
    if kind == "icon" and icon:
        png = icon_png(b, icon, colour)
        if png:
            import deck_theme
            b.shape(s, MSO_SHAPE.OVAL, (l, t, w, h), deck_theme.mix(colour, b.slide_bg, 0.85))
            pad = int(w * 0.22)
            s.shapes.add_picture(str(png), l + pad, t + pad, w - 2 * pad, h - 2 * pad)
            return
        kind = "number"
    if kind in ("check", "cross"):
        png = icon_png(b, "circle-check" if kind == "check" else "circle-x", colour)
        if png:
            s.shapes.add_picture(str(png), l, t, w, h)
            return
        kind = "dot"
    if kind == "numeral":
        b.text(s, (l, t, w, h), [(0, [(f"{i + 1:02d}", "b")])], 36, 0, {"title": "", "line": 0},
               colour=colour, family="+mj-lt", anchor=MSO_ANCHOR.MIDDLE, min_size=20, what=" numeral")
        return
    if kind == "number":
        c = b.shape(s, MSO_SHAPE.OVAL, (l, t, w, h), colour)
        tf = c.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        b.text(s, None, [(0, [(str(i + 1), "b")])], 20, 0, {"title": "", "line": 0}, colour="FFFFFF",
               family="+mj-lt", align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=c, min_size=10,
               what=" number")
        return
    d = int(min(w, h) * 0.32)                                  # dot
    b.shape(s, MSO_SHAPE.OVAL, (l + (w - d) // 2, t + (h - d) // 2, d, d), colour)


def rows(b, s, box, items, n, sl, kind="dot", size=22, dividers=True, colour=None, gap_in=0.16, top=False,
         row_in=1.45):
    """Items stacked down `box`, each a marker + label/text; one text size for all of them."""
    import deck
    l, t, w, h = (int(v) for v in box)
    k = len(items)
    gap = int(gap_in * IN)
    rh = min((h - gap * (k - 1)) // k, int(row_in * IN))
    y0 = t if top else t + (h - (k * rh + (k - 1) * gap)) // 2
    if any(it[0] for it in items):
        kind = "icon"
    mk = int(0.95 * IN) if kind == "numeral" else int(0.6 * IN) if kind in ("icon", "number") else \
        int(0.42 * IN) if kind in ("check", "cross") else int(0.3 * IN)
    ink = b.ink()
    lv = {0: {"size": size, "marL": 0, "spc": 0, "ln": 1.0}, 1: {"size": size - 3, "marL": 0, "spc": 3, "ln": 1.0}}
    filled = []
    for i, (icon, head, body, subs) in enumerate(items):
        y = y0 + i * (rh + gap)
        m = min(mk, rh)
        marker(b, s, "icon" if icon else kind, (l, y + (rh - m) // 2, mk if kind == "numeral" else m, m), i, icon,
               colour)
        paras = item_paras(b, head, body, subs)
        tx = l + mk + int(0.22 * IN)
        tb = s.shapes.add_textbox(tx, y, w - (tx - l), rh)
        args = ([(lvl, rs, False) for lvl, rs in paras], b.tpl, b.problems, b.where(n, sl))
        kw = dict(min_size=14, levels=lv, colour=ink["text"], anchor=MSO_ANCHOR.MIDDLE)
        filled.append((tb, args, kw, deck.fill_text(tb, *args, **kw)))
        if head:                          # the label in the title colour
            for r in tb.text_frame.paragraphs[0].runs:
                r.font.color.rgb = rgb(ink["title"])
        if dividers and i < k - 1:
            b.shape(s, MSO_SHAPE.RECTANGLE, (tx, y + rh + gap // 2, w - (tx - l), Pt(0.75)), ink["line"])
    b.uniform(filled)
    for tb, *_ in filled:
        for r in tb.text_frame.paragraphs[0].runs:
            if r.font.bold:
                r.font.color.rgb = rgb(ink["title"])


def rgb(h):
    from pptx.dml.color import RGBColor
    return RGBColor.from_string(h)


def blocks_by_heading(sl):
    """### headings split a slide into groups: [(heading, classes, items, paras)] and the intro blocks."""
    intro, groups = [], []
    for blk in sl["blocks"]:
        if blk["kind"] == "heading":
            groups.append([blk["text"], blk.get("classes", []), [], []])
        elif groups:
            if blk["kind"] == "bullets":
                for lvl, t in blk["items"]:
                    if lvl == 0 or not groups[-1][2]:
                        groups[-1][2].append([t, []])
                    else:
                        groups[-1][2][-1][1].append(t)
            else:
                groups[-1][3].append(blk)
        else:
            intro.append(blk)
    return intro, groups


def intro_line(b, s, intro, box, n, sl):
    """A sentence above a design; → the box left under it."""
    import deck
    paras = deck.plain(b.paras_of(intro))
    l, t, w, h = box
    if not paras:
        return box
    th = int(0.75 * IN)
    tb = s.shapes.add_textbox(l, t, w, th)
    deck.fill_text(tb, paras, b.tpl, b.problems, b.where(n, sl), min_size=16,
                   levels={0: {"size": 20, "marL": 0, "spc": 0, "ln": 1.0}}, colour=b.ink()["muted"],
                   anchor=MSO_ANCHOR.TOP)
    return (l, t + th + int(0.1 * IN), w, h - th - int(0.1 * IN))


# ── designs ──────────────────────────────────────────────────
def lead(b, n, sl, cards=False):
    """The problem statement: one sentence large, its 2–4 points beside it (or as cards below)."""
    import deck
    items, rest = b.items_of(sl)
    sentence = " ".join(x["text"] for x in rest if x["kind"] == "para")
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = b.tpl.content_box
    ink, P = b.ink(), b.tpl.palette
    its = items_with_icons(b, items)
    if cards:
        sh = int(1.35 * IN)
        b.text(s, (l, t, w, sh), [(0, deck.runs(sentence, b.cites))], 28, n, sl, colour=ink["title"],
               family="+mj-lt", anchor=MSO_ANCHOR.MIDDLE, min_size=20, what=" statement")
        card_row(b, s, (l, t + sh + int(0.3 * IN), w, h - sh - int(0.3 * IN)), its, n, sl)
    else:
        lw = int(w * 0.44)
        b.shape(s, MSO_SHAPE.RECTANGLE, (l, t + int(h * 0.2), int(0.07 * IN), int(h * 0.6)), ink["accent"])
        b.text(s, (l + int(0.3 * IN), t, lw - int(0.3 * IN), h), [(0, deck.runs(sentence, b.cites))], 30, n, sl,
               colour=ink["title"], family="+mj-lt", anchor=MSO_ANCHOR.MIDDLE, min_size=20, what=" statement")
        rows(b, s, (l + lw + int(0.5 * IN), t, w - lw - int(0.5 * IN), h), its, n, sl, kind="dot")
    b.finish(s, "title_only", sl["notes"])


def card_row(b, s, box, its, n, sl):
    """Items as equal cards in one row: icon or number, label, text."""
    import deck
    l, t, w, h = (int(v) for v in box)
    ink = b.ink()
    k = len(its)
    gap = int(0.3 * IN)
    cw = (w - gap * (k - 1)) // k
    pad, mk = int(0.25 * IN), int(0.55 * IN)
    lv = {0: {"size": 22, "marL": 0, "spc": 0, "ln": 1.0}, 1: {"size": 18, "marL": 0, "spc": 6, "ln": 1.0}}
    need = 0
    for icon, head, body, subs in its:
        paras = [(lvl, rs, False) for lvl, rs in item_paras(b, head, body, subs)]
        need = max(need, deck.text_height(paras, b.tpl, "+mn-lt", 1.0, (cw - 2 * pad) / 12700 - 14.4, lv) / 0.9)
    ch = min(h, int(need * 12700 + mk + 3 * pad + 0.2 * IN))
    filled = []
    for i, (icon, head, body, subs) in enumerate(its):
        x = l + i * (cw + gap)
        card = b.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, (x, t, cw, ch), ink["card"])
        card.adjustments[0] = 0.06
        marker(b, s, "icon" if icon else "number", (x + pad, t + pad, mk, mk), i, icon)
        tb = s.shapes.add_textbox(x + pad, t + pad + mk + int(0.12 * IN), cw - 2 * pad, ch - 2 * pad - mk - int(0.12 * IN))
        paras = [(lvl, rs, False) for lvl, rs in item_paras(b, head, body, subs)]
        args = (paras, b.tpl, b.problems, b.where(n, sl) + " card")
        kw = dict(min_size=14, levels=lv, colour=ink["text"], anchor=MSO_ANCHOR.TOP)
        filled.append((tb, args, kw, deck.fill_text(tb, *args, **kw)))
    b.uniform(filled)
    for tb, *_ in filled:
        for r in tb.text_frame.paragraphs[0].runs:
            if r.font.bold:
                r.font.color.rgb = rgb(ink["title"])


def numbered(b, n, sl, check=False):
    """Objectives / contributions: big numerals (or ✓) down the slide, thin rules between."""
    items, rest = b.items_of(sl)
    s = b.new("title_only")
    title(b, s, sl, n)
    box = intro_line(b, s, rest, b.tpl.content_box, n, sl)
    its = items_with_icons(b, items) if check else [parse_item(t) + (subs,) for t, subs in items]
    if check:
        its = [(None, h, bd, sb) for _, h, bd, sb in its]
    rows(b, s, box, its, n, sl, kind="check" if check else "numeral", size=22)
    b.finish(s, "title_only", sl["notes"])


def compare(b, n, sl):
    """Two options side by side with "vs" between; the better one ({.win} on its ### heading, or
    named proposed/ours) is outlined in the accent colour, the other is quieter."""
    import deck
    intro, groups = blocks_by_heading(sl)
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = intro_line(b, s, intro, b.tpl.content_box, n, sl)
    ink, P = b.ink(), b.tpl.palette
    win = next((k for k, g in enumerate(groups) if "win" in g[1]), None)
    if win is None:
        win = next((k for k, g in enumerate(groups)
                    if re.search(r"\b(proposed|ours|our|new)\b", g[0], re.I)), None)
    gap = int(0.9 * IN)
    cw = (w - gap) // 2
    hb = int(0.85 * IN)
    lv = {0: {"size": 20, "marL": 0, "spc": 8, "ln": 1.0}}
    need = 0
    for head, _, items, paras in groups[:2]:
        lines = [(0, [("•  ", "")] + deck.runs(x2, b.cites), False) for it, subs in items for x2 in [it] + subs]
        need = max(need, deck.text_height(lines, b.tpl, "+mn-lt", 1.0, (cw - 0.6 * IN) / 12700 - 14.4, lv) / 0.9)
    ch = min(h, max(int(need * 12700) + hb + int(0.9 * IN), int(h * 0.6)))
    t += (h - ch) // 3
    h = ch
    filled = []
    for k, (head, _, items, paras) in enumerate(groups[:2]):
        x = l + k * (cw + gap)
        winner = k == win
        card = b.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, (x, t, cw, h), "FFFFFF" if winner and not b.dark
                       else ink["card"])
        card.adjustments[0] = 0.04
        if winner:
            card.line.color.rgb = rgb(ink["accent"])
            card.line.width = Pt(2.5)
        b.text(s, (x + int(0.3 * IN), t + int(0.15 * IN), cw - int(0.6 * IN), hb),
               [(0, [(x2, st + "b") for x2, st in deck.runs(head, b.cites)])], 24, n, sl,
               colour=ink["accent"] if winner else ink["muted"], family="+mj-lt", anchor=MSO_ANCHOR.MIDDLE,
               min_size=16, what=" heading")
        b.shape(s, MSO_SHAPE.RECTANGLE, (x + int(0.3 * IN), t + hb + int(0.15 * IN), cw - int(0.6 * IN), Pt(1)),
                ink["accent"] if winner else ink["line"])
        lines = [(0, deck.runs(x2, b.cites), True) for it, subs in items for x2 in [it] + subs]
        lines += [(0, r, False) for _, r, _ in b.paras_of(paras)]
        tb = s.shapes.add_textbox(x + int(0.3 * IN), t + hb + int(0.35 * IN), cw - int(0.6 * IN),
                                  h - hb - int(0.55 * IN))
        args = ([(0, [("•  ", "")] + rs if bul else rs, False) for _, rs, bul in lines], b.tpl, b.problems,
                b.where(n, sl))
        kw = dict(min_size=14, levels=lv, colour=ink["text"] if winner or win is None else ink["muted"],
                  anchor=MSO_ANCHOR.TOP)
        filled.append((tb, args, kw, deck.fill_text(tb, *args, **kw)))
    b.uniform(filled)
    d = int(0.75 * IN)
    vs = b.shape(s, MSO_SHAPE.OVAL, (l + cw + (gap - d) // 2, t + (h - d) // 2, d, d),
                 ink["accent"] if b.dark else P["dk2"])
    tf = vs.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    b.text(s, None, [(0, [("vs", "b")])], 20, n, sl, colour="FFFFFF", family="+mj-lt", align=PP_ALIGN.CENTER,
           anchor=MSO_ANCHOR.MIDDLE, shape=vs, min_size=10, what=" vs")
    if len(groups) > 2:
        b.problems.append(("WARN", f"{b.where(n, sl)}: a comparison shows two ### groups; the rest is left out"))
    b.finish(s, "title_only", sl["notes"])


PRO = r"advantage|pro\b|pros\b|strength|benefit|merit|why it works|good"
CON = r"limitation|disadvantage|con\b|cons\b|weakness|drawback|challenge|risk|issue|problem|bad"


def proscons(b, n, sl):
    """Advantages / limitations: ✓ and ✗ columns."""
    import deck_theme
    intro, groups = blocks_by_heading(sl)
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = intro_line(b, s, intro, b.tpl.content_box, n, sl)
    ink, P = b.ink(), b.tpl.palette
    gap = int(0.6 * IN)
    cw = (w - gap) // 2
    for k, (head, _, items, _) in enumerate(groups[:2]):
        con = bool(re.search(CON, head, re.I)) and not re.search(PRO, head, re.I)
        colour = deck_theme.mix(ink["text"], "FFFFFF" if not b.dark else "000000", 0.35) if con else ink["accent"]
        x = l + k * (cw + gap)
        b.text(s, (x, t, cw, int(0.7 * IN)), [(0, [(head, "b")])], 24, n, sl, colour=colour, family="+mj-lt",
               anchor=MSO_ANCHOR.MIDDLE, min_size=16, what=" heading")
        b.shape(s, MSO_SHAPE.RECTANGLE, (x, t + int(0.75 * IN), cw, Pt(2)), colour)
        its = [parse_item(it) + (subs,) for it, subs in items]
        rows(b, s, (x, t + int(1.1 * IN), cw, h - int(1.1 * IN)), [(None, hh, bd, sb) for _, hh, bd, sb in its],
             n, sl, kind="cross" if con else "check", size=20, dividers=False, colour=colour, top=True, row_in=0.85)
    b.finish(s, "title_only", sl["notes"])


def columns(b, n, sl):
    """Three or four ### groups side by side (e.g. one paper per column: method, result, gap)."""
    import deck
    intro, groups = blocks_by_heading(sl)
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = intro_line(b, s, intro, b.tpl.content_box, n, sl)
    ink, P = b.ink(), b.tpl.palette
    k = min(len(groups), 4)
    gap = int(0.3 * IN)
    cw = (w - gap * (k - 1)) // k
    lv = {0: {"size": 18, "marL": 0, "spc": 8, "ln": 1.0}}
    need = 0
    for head, cls, items, paras in groups[:4]:
        ls = [(0, deck.runs(x2, b.cites), False) for it, subs in items for x2 in [it] + subs]
        need = max(need, deck.text_height(ls, b.tpl, "+mn-lt", 1.0, (cw - 0.5 * IN) / 12700 - 14.4, lv) / 0.9)
    ch = min(h, max(int(need * 12700) + int(1.7 * IN), int(h * 0.55)))
    t += (h - ch) // 3
    h = ch
    filled = []
    for i, (head, cls, items, paras) in enumerate(groups[:4]):
        x = l + i * (cw + gap)
        hi = "win" in cls
        card = b.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, (x, t, cw, h), ink["card"])
        card.adjustments[0] = 0.05
        if hi:
            card.line.color.rgb = rgb(ink["accent"])
            card.line.width = Pt(2.5)
        b.shape(s, MSO_SHAPE.RECTANGLE, (x + int(0.25 * IN), t + int(0.25 * IN), int(0.6 * IN), Pt(3)), ink["accent"])
        b.text(s, (x + int(0.25 * IN), t + int(0.35 * IN), cw - int(0.5 * IN), int(0.8 * IN)),
               [(0, [(x2, st + "b") for x2, st in deck.runs(head, b.cites)])], 22, n, sl, colour=ink["title"],
               family="+mj-lt", anchor=MSO_ANCHOR.TOP, min_size=14, what=" heading")
        lines = []
        for it, subs in items:
            ic, hh, bd = parse_item(it)
            lines.append((0, ([(hh + ": ", "b")] if hh else []) + deck.runs(bd, b.cites)))
            lines += [(0, deck.runs(x2, b.cites)) for x2 in subs]
        lines += [(0, r) for _, r, _ in b.paras_of(paras)]
        tb = s.shapes.add_textbox(x + int(0.25 * IN), t + int(1.2 * IN), cw - int(0.5 * IN), h - int(1.4 * IN))
        args = ([(0, r, False) for _, r in lines], b.tpl, b.problems, b.where(n, sl))
        kw = dict(min_size=14, levels={0: {"size": 18, "marL": 0, "spc": 8, "ln": 1.0}}, colour=ink["text"],
                  anchor=MSO_ANCHOR.TOP)
        filled.append((tb, args, kw, deck.fill_text(tb, *args, **kw)))
    b.uniform(filled)
    b.finish(s, "title_only", sl["notes"])


def callouts(b, n, sl):
    """A diagram large, with its 2–4 notes numbered beside it (or in a row under a wide diagram)."""
    items, rest = b.items_of(sl)
    pic = next(x for x in rest if x["kind"] == "image")
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = b.tpl.content_box
    its = [(None, hh, bd, sb) for _, hh, bd, sb in items_with_icons(b, items)]
    if b.wide(pic):
        ph = int(h * 0.56)
        b.picture_block(s, (l, t, w, ph), pic, n, sl)
        k = len(its)
        gap = int(0.3 * IN)
        cw = (w - gap * (k - 1)) // k
        for i, it in enumerate(its):
            rows(b, s, (l + i * (cw + gap), t + ph + int(0.25 * IN), cw, h - ph - int(0.25 * IN)), [it], n, sl,
                 kind="number", size=18, dividers=False)
            # rows() numbers each call from 1: renumber
            _renumber_last(s, i + 1)
    else:
        pw = int(w * 0.6)
        b.picture_block(s, (l, t, pw, h), pic, n, sl)
        rows(b, s, (l + pw + int(0.4 * IN), t, w - pw - int(0.4 * IN), h), its, n, sl, kind="number", size=20)
    b.finish(s, "title_only", sl["notes"])


def _renumber_last(s, k):
    for sh in reversed(list(s.shapes)):
        if sh.has_text_frame and sh.text_frame.text.strip().isdigit() and sh.width < 0.7 * IN:
            sh.text_frame.paragraphs[0].runs[0].text = str(k)
            return


def insight(b, n, sl):
    """A chart or figure with the one thing to see in it, in a takeaway box beside it."""
    import deck
    pic = next(x for x in sl["blocks"] if x["kind"] == "image")
    text = " ".join(x["text"] for x in sl["blocks"] if x["kind"] == "para")
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = b.tpl.content_box
    ink, P = b.ink(), b.tpl.palette
    pw = int(w * 0.64)
    b.picture_block(s, (l, t, pw, h), pic, n, sl)
    bx, bw = l + pw + int(0.35 * IN), w - pw - int(0.35 * IN)
    bh = int(h * 0.62)
    by = t + (h - bh) // 2
    b.shape(s, MSO_SHAPE.RECTANGLE, (bx, by, bw, bh), ink["card"])
    b.shape(s, MSO_SHAPE.RECTANGLE, (bx, by, int(0.08 * IN), bh), ink["accent"])
    b.text(s, (bx + int(0.3 * IN), by + int(0.2 * IN), bw - int(0.5 * IN), int(0.45 * IN)),
           [(0, [("Takeaway", "b")])], 16, n, sl, colour=ink["accent"], min_size=12, what=" label")
    b.text(s, (bx + int(0.3 * IN), by + int(0.7 * IN), bw - int(0.5 * IN), bh - int(0.9 * IN)),
           [(0, deck.runs(text, b.cites))], 24, n, sl, colour=ink["title"], family="+mj-lt", min_size=16,
           anchor=MSO_ANCHOR.TOP, what=" takeaway")
    b.finish(s, "title_only", sl["notes"])


def numbers_chart(b, n, sl):
    """Key results: the numbers stacked on the left, the chart that shows them on the right."""
    import deck
    pic = next(x for x in sl["blocks"] if x["kind"] == "image")
    cards = [[x.strip() for x in t.split("::", 1)] for blk in sl["blocks"] if blk["kind"] == "bullets"
             for _, t in blk["items"] if "::" in t]
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = b.tpl.content_box
    ink, P = b.ink(), b.tpl.palette
    lw = int(w * 0.3)
    k = len(cards)
    ch = h // max(k, 1)
    for i, (v, lab) in enumerate(cards):
        y = t + i * ch
        b.text(s, (l, y + int(ch * 0.08), lw, int(ch * 0.55)), [(0, [(v, "b")])], 48 if i == 0 else 40, n, sl,
               colour=ink["accent"] if i == 0 else ink["title"], family="+mj-lt", anchor=MSO_ANCHOR.BOTTOM,
               min_size=24, what=" number")
        b.text(s, (l, y + int(ch * 0.63), lw, int(ch * 0.35)), [(0, deck.runs(lab, b.cites))], 18, n, sl,
               colour=ink["muted"], min_size=12, what=" label")
    b.shape(s, MSO_SHAPE.RECTANGLE, (l + lw + int(0.2 * IN), t + int(0.2 * IN), Pt(1), h - int(0.4 * IN)),
            ink["line"])
    b.picture_block(s, (l + lw + int(0.5 * IN), t, w - lw - int(0.5 * IN), h), pic, n, sl)
    b.finish(s, "title_only", sl["notes"])


def definition(b, n, sl):
    """A key term large, its meaning under it: '**Non-IID data**: clients see different …'."""
    import deck
    items, rest = b.items_of(sl)
    text = items[0][0] if items else " ".join(x["text"] for x in rest if x["kind"] == "para")
    _, term, meaning = parse_item(text)
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = b.tpl.content_box
    ink, P = b.ink(), b.tpl.palette
    y = t + int(h * 0.18)
    b.text(s, (l, y, w, int(1.3 * IN)), [(0, [(term or sl["title"], "b")])], 54, n, sl, colour=ink["accent"],
           family="+mj-lt", anchor=MSO_ANCHOR.BOTTOM, min_size=28, what=" term")
    b.shape(s, MSO_SHAPE.RECTANGLE, (l, y + int(1.45 * IN), int(1.2 * IN), int(0.06 * IN)), P["accent2"])
    b.text(s, (l, y + int(1.7 * IN), int(w * 0.8), h - int(h * 0.18) - int(1.7 * IN)),
           [(0, deck.runs(meaning, b.cites))] + [(0, deck.runs(x, b.cites)) for x in (items[0][1] if items else [])],
           26, n, sl, colour=ink["text"], min_size=18, anchor=MSO_ANCHOR.TOP, what=" definition")
    b.finish(s, "title_only", sl["notes"])


def takeaways(b, n, sl):
    """Conclusion: 2–4 points as columns under big numerals; no boxes."""
    import deck
    items, rest = b.items_of(sl)
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = intro_line(b, s, rest, b.tpl.content_box, n, sl)
    ink, P = b.ink(), b.tpl.palette
    k = len(items)
    gap = int(0.45 * IN)
    cw = (w - gap * (k - 1)) // k
    lv = {0: {"size": 22, "marL": 0, "spc": 0, "ln": 1.0}, 1: {"size": 18, "marL": 0, "spc": 6, "ln": 1.0}}
    need = max(deck.text_height([(lvl, rs, False) for lvl, rs in item_paras(b, *parse_item(text)[1:], subs)],
                                b.tpl, "+mn-lt", 1.0, cw / 12700 - 14.4, lv) / 0.9 for text, subs in items)
    group = int(1.4 * IN + need * 12700)
    top = t + max(0, (h - group) // 3)
    filled = []
    for i, (text, subs) in enumerate(items):
        icon, head, body = parse_item(text)
        x = l + i * (cw + gap)
        b.text(s, (x, top, cw, int(1.1 * IN)), [(0, [(f"{i + 1}", "b")])], 66, n, sl, colour=ink["accent"],
               family="+mj-lt", anchor=MSO_ANCHOR.BOTTOM, min_size=40, what=" numeral")
        b.shape(s, MSO_SHAPE.RECTANGLE, (x, top + int(1.2 * IN), cw, Pt(1.5)), ink["line"])
        tb = s.shapes.add_textbox(x, top + int(1.4 * IN), cw, t + h - top - int(1.4 * IN))
        paras = [(lvl, rs, False) for lvl, rs in item_paras(b, head, body, subs)]
        lv = {0: {"size": 22, "marL": 0, "spc": 0, "ln": 1.0}, 1: {"size": 18, "marL": 0, "spc": 6, "ln": 1.0}}
        args = (paras, b.tpl, b.problems, b.where(n, sl))
        kw = dict(min_size=14, levels=lv, colour=ink["text"], anchor=MSO_ANCHOR.TOP)
        filled.append((tb, args, kw, deck.fill_text(tb, *args, **kw)))
    b.uniform(filled)
    for tb, *_ in filled:
        for r in tb.text_frame.paragraphs[0].runs:
            if r.font.bold:
                r.font.color.rgb = rgb(ink["title"])
    b.finish(s, "title_only", sl["notes"])


def agenda_slide(b, n, sections):
    """Outline after the cover: the sections, numbered."""
    sl = {"title": b.cfg.get("agenda_title", "Outline"), "classes": [], "attrs": {}, "blocks": [], "notes": [],
          "line": 0}
    b.cur = sl
    s = b.new("title_only")
    title(b, s, sl, n)
    l, t, w, h = b.tpl.content_box
    rows(b, s, (l, t, int(w * 0.7), h), [(None, x, "", []) for x in sections], n, sl, kind="numeral", size=24)
    b.finish(s, "title_only", [])
    b.cur = None


def progress(b, s, sections, current):
    """On a section divider: every section listed, this one highlighted (where we are in the talk)."""
    import deck_theme
    W, H = b.tpl.W, b.tpl.H
    t = s.shapes.title
    t.left, t.top, t.width, t.height = t.left, t.top, int(W * 0.5) - t.left, t.height
    ink = b.ink()
    x, w = int(W * 0.58), int(W * 0.36)
    k = len(sections)
    rh = min(int(0.62 * IN), int(H * 0.7) // k)
    y0 = (H - rh * k) // 2
    dim = deck_theme.mix(ink["title"], b.slide_bg, 0.55)
    b.shape(s, MSO_SHAPE.RECTANGLE, (x - int(0.35 * IN), y0, Pt(1.5), rh * k), dim)
    for i, name in enumerate(sections):
        cur = i == current
        b.text(s, (x, y0 + i * rh, w, rh), [(0, [(f"{i + 1:02d}   ", "b"), (name, "b" if cur else "")])],
               22 if cur else 18, 0, {"title": "section list", "line": 0},
               colour=ink["title"] if cur else dim, family="+mj-lt", anchor=MSO_ANCHOR.MIDDLE, min_size=12)
        if cur:
            b.shape(s, MSO_SHAPE.RECTANGLE, (x - int(0.38 * IN), y0 + i * rh + rh // 4, Pt(4), rh // 2),
                    b.tpl.palette["accent2"])


# ── choosing ─────────────────────────────────────────────────
PROCESS = r"\b(how|pipeline|process|workflow|steps?|algorithm|procedure|method(ology)?|approach|flow|stages?)\b"
CONCLUDE = r"\b(conclusions?|summary|takeaways?|contributions?|key points|in short)\b"
FUTURE = r"\b(future|roadmap|next steps?|plans?|extensions?)\b"
OBJECTIVE = r"\b(objectives?|goals?|aims?|scope)\b"


def choose(sl, kind):
    """→ (design, why) for a slide without an explicit design class; design None = keep `kind`."""
    blocks = sl["blocks"]
    kinds = [x["kind"] for x in blocks]
    title = sl["title"]
    lists = [x for x in blocks if x["kind"] == "bullets"]
    paras = [x for x in blocks if x["kind"] == "para"]
    imgs = [x for x in blocks if x["kind"] == "image"]
    tops = [t for x in lists for lvl, t in x["items"] if lvl == 0]
    if kind not in ("bullets", "two-col", "image", "results"):
        return None, ""
    if "heading" in kinds:
        groups = blocks_by_heading(sl)[1]
        if len(groups) == 2:
            if any(re.search(PRO, g[0], re.I) for g in groups) and any(re.search(CON, g[0], re.I) for g in groups):
                return "proscons", "advantages and limitations under ### headings"
            return "compare", "two groups under ### headings"
        if len(groups) >= 3:
            return "columns", f"{len(groups)} groups under ### headings"
        return None, ""
    if kind == "results" and imgs:
        return "numbers_chart", "numbers and a chart"
    if "split" in kinds or "table" in kinds or "equation" in kinds:
        return None, ""
    if len(imgs) == 1 and not any("eq" in x.get("classes", []) for x in imgs):
        if lists and not paras and 2 <= len(tops) <= 4 and all(len(t.split()) <= 18 for t in tops):
            return "callouts", "a picture with a few short notes"
        if paras and not lists and sum(len(p["text"].split()) for p in paras) >= 6:
            return "insight", "a picture with one message about it"
        return None, ""
    if imgs or len(lists) != 1:
        return None, ""
    items = lists[0]["items"]
    has_sub = any(lvl for lvl, _ in items)
    if len(paras) == 1 and 2 <= len(tops) <= 4 and len(paras[0]["text"].split()) <= 40 and not has_sub \
            and kinds.index("para") < kinds.index("bullets"):
        return "lead", "a statement and its points"
    if paras and len(paras) > 1:
        return None, ""
    if len(tops) == 1 and parse_item(tops[0])[1] and not paras:
        return "definition", "one term and its meaning"
    if re.search(CONCLUDE, title, re.I) and 2 <= len(tops) <= 4 and not has_sub:
        return "takeaways", "a conclusion in a few points"
    if re.search(FUTURE, title, re.I) and 2 <= len(tops) <= 6 and not has_sub:
        return "roadmap", "future work as a sequence"
    if lists[0].get("ordered") and 2 <= len(tops) <= 6 and not has_sub:
        if re.search(PROCESS, title, re.I):
            return "steps", "an ordered list of process steps"
        return "numbered", "an ordered list"
    if re.search(OBJECTIVE, title, re.I) and 2 <= len(tops) <= 6 and not has_sub:
        return "numbered", "objectives"
    return None, ""
