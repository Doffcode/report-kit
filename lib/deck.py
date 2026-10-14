"""Slide decks: slides/slides.md → slides/out/slides.pptx (+ .pdf, _build/slides.png).

Content goes into the layouts of a real .pptx (the kit theme or your college
template), so every title, bullet, table and chart stays a native, editable
PowerPoint object. Layout is the build's job: text is measured with the real
font metrics and sized to fit (autofit is switched off so PowerPoint and
LibreOffice don't shrink it again), pictures are scaled to fit without cropping,
and text that still doesn't fit fails the build instead of spilling off a slide.
"""
import copy
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import yaml
from lxml import etree
from PIL import Image, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Pt

KIT = Path(__file__).resolve().parent.parent
EMU_PT = 12700
SAFETY = 0.90            # use only 90% of a box's height: viewers wrap slightly differently
MIN_BODY, MIN_TITLE, MIN_SMALL = 18, 24, 12
METRIC_TWINS = {         # fonts with identical letter widths: measuring with one is measuring the other
    "calibri": {"carlito"}, "cambria": {"caladea"}, "arial": {"liberation sans", "arimo"},
    "times new roman": {"liberation serif", "tinos"}, "courier new": {"liberation mono", "cousine"},
    "helvetica": {"liberation sans", "arimo"},
}
CODE_FONT = "Courier New"


def say(msg):
    print(f"\033[1m▸\033[0m {msg}", flush=True)


# ── project files ────────────────────────────────────────────
def find_slides_dir(start: Path) -> Path:
    for d in (start, start / "slides"):
        if (d / "slides.yml").exists():
            return d
    sys.exit("No slides.yml here or in ./slides. Run `report slides init` first.")


def load_config(sdir: Path) -> dict:
    cfg = yaml.safe_load((sdir / "slides.yml").read_text()) or {}
    rdir = sdir.parent / "report"
    me = {}
    for p in (KIT / "me.yml", rdir / "me.yml", sdir / "me.yml"):   # later files override earlier ones
        if p.exists():
            d = (yaml.safe_load(p.read_text()) or {}).get("me") or {}
            if d.get("logo"):
                d["logo"] = str((p.parent / d["logo"]).resolve())
            me = {**me, **d}
    cfg["me"] = {**me, **(cfg.get("me") or {})}
    cfg["_report"] = rdir if (rdir / "report.yml").exists() else None
    if not cfg.get("title") and cfg["_report"]:
        cfg["title"] = (yaml.safe_load((rdir / "report.yml").read_text()) or {}).get("title")
    cfg.setdefault("title", "Untitled talk")
    cfg.setdefault("date", __import__("time").strftime("%B %Y"))
    return cfg


def template_path(sdir: Path, cfg: dict) -> Path:
    """A .pptx template (college/company) if slides.yml names one, else a kit theme made from `theme:`."""
    t = cfg.get("template") or "kit"
    if t == "kit":
        import deck_theme
        return deck_theme.make(deck_theme.spec_of(cfg), sdir / "_build")
    for p in (sdir / t, KIT / "templates/slides" / f"{t}.pptx", KIT / "templates/slides" / t):
        if p.is_file():
            return p
    sys.exit(f"template {t!r}: not found (looked in slides/ and the kit's templates/slides/)")


# ── slides.md ────────────────────────────────────────────────
HEAD = re.compile(r"^(#{1,2})\s+(.*?)\s*(\{([^}]*)\})?\s*$")
LIST = re.compile(r"^(\s*)([-*+]|\d+[.)])\s+(.*)$")
IMAGE = re.compile(r"^!\[(.*?)\]\((.*?)\)\s*(\{([^}]*)\})?\s*$")


def parse(text: str):
    """→ list of slides: {title, classes, level (1 section / 2 slide), blocks, notes, line}."""
    slides, cur, in_notes, para = [], None, False, []

    def flush_para():
        if para and cur is not None:
            cur["blocks"].append({"kind": "para", "text": " ".join(s.strip() for s in para)})
        para.clear()

    in_eq = None
    for n, line in enumerate(text.splitlines(), 1):
        if in_eq is not None:
            if line.strip().endswith("$$"):
                in_eq.append(line.strip()[:-2])
                cur["blocks"].append({"kind": "equation", "tex": " ".join(in_eq).strip()})
                in_eq = None
            else:
                in_eq.append(line.strip())
            continue
        if in_notes:
            if line.strip() == ":::":
                in_notes = False
            else:
                cur["notes"].append(line)
            continue
        m = HEAD.match(line)
        if m:
            flush_para()
            classes = re.findall(r"(?<![\w\"])\.([\w-]+)", m.group(4) or "")
            attrs = {k: v.strip('"') for k, v in re.findall(r'([\w-]+)=("[^"]*"|[^\s}]+)', m.group(4) or "")}
            cur = {"title": m.group(2), "classes": classes, "attrs": attrs, "level": len(m.group(1)),
                   "blocks": [], "notes": [], "line": n}
            slides.append(cur)
            continue
        if cur is None:
            continue
        s = line.strip()
        if s.startswith(":::") and "notes" in s:
            flush_para()
            in_notes = True
            continue
        if not s:
            flush_para()
            continue
        if s == "|||":
            flush_para()
            cur["blocks"].append({"kind": "split"})
            continue
        if s.startswith("$$") and (len(s) == 2 or not s.endswith("$$")):
            flush_para()
            in_eq = [s[2:]]
            continue
        m = re.fullmatch(r"\$\$(.+)\$\$", s)
        if m:
            flush_para()
            cur["blocks"].append({"kind": "equation", "tex": m.group(1).strip()})
            continue
        m = IMAGE.match(s)
        if m:
            flush_para()
            attrs = m.group(4) or ""
            cur["blocks"].append({"kind": "image", "caption": m.group(1), "path": m.group(2),
                                  "classes": re.findall(r"\.([\w-]+)", attrs),
                                  "attrs": {k: v.strip('"') for k, v in
                                            re.findall(r'([\w-]+)=("[^"]*"|[^\s}]+)', attrs)}})
            continue
        if s.startswith("|"):
            flush_para()
            blk = cur["blocks"][-1] if cur["blocks"] and cur["blocks"][-1]["kind"] == "table" else None
            if blk is None:
                blk = {"kind": "table", "rows": [], "align": None}
                cur["blocks"].append(blk)
            cells = [c.strip() for c in s.strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                blk["align"] = ["r" if c.endswith(":") and not c.startswith(":") else
                                "c" if c.startswith(":") and c.endswith(":") else "l" for c in cells]
            else:
                blk["rows"].append(cells)
            continue
        m = re.match(r"^###\s+(.*?)\s*(\{([^}]*)\})?\s*$", s)
        if m:                                              # a group inside a slide (compare, columns)
            flush_para()
            cur["blocks"].append({"kind": "heading", "text": m.group(1),
                                  "classes": re.findall(r"\.([\w-]+)", m.group(3) or "")})
            continue
        m = LIST.match(line)
        if m:
            flush_para()
            blk = cur["blocks"][-1] if cur["blocks"] and cur["blocks"][-1]["kind"] == "bullets" else None
            if blk is None:
                blk = {"kind": "bullets", "items": [], "ordered": m.group(2)[0].isdigit()}
                cur["blocks"].append(blk)
            blk["items"].append((len(m.group(1).expandtabs(2)) // 2, m.group(3)))
            continue
        if cur["blocks"] and cur["blocks"][-1]["kind"] == "bullets" and line.startswith("  "):
            lvl, t = cur["blocks"][-1]["items"][-1]            # wrapped continuation of a bullet
            cur["blocks"][-1]["items"][-1] = (lvl, t + " " + s)
            continue
        para.append(s)
    flush_para()
    return slides


TYPES = ("results", "references", "qa", "two-col", "image", "table", "bullets", "title", "cards", "steps",
         "timeline", "statement", "quote", "image-left", "image-right", "agenda", "lead", "lead-cards", "numbered",
         "checklist", "compare", "proscons", "columns", "callouts", "insight", "definition", "takeaways", "roadmap",
         "numbers-chart")


def slide_kind(sl):
    for k in TYPES:
        if k in sl["classes"]:
            return k
    kinds = [b["kind"] for b in sl["blocks"]]
    if sl["level"] == 1:
        return "section"
    if "split" in kinds:
        return "two-col"
    if "image" in kinds and any(k in ("bullets", "para", "table") for k in kinds if k != "image") \
            and not all(k in ("image", "para") for k in kinds):
        return "two-col"            # text + picture side by side
    if "image" in kinds:
        return "image"
    if "table" in kinds:
        return "table"
    return "bullets"


# ── inline text → runs ───────────────────────────────────────
INLINE = re.compile(r"(\[\[TODO[^\]]*\]\])|\*\*(.+?)\*\*|\*(.+?)\*|(?<!\w)_(.+?)_(?!\w)|`([^`]+)`|\[(@[^\]]+)\]|\$([^$]+)\$")


GREEK = {"alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε", "eta": "η", "theta": "θ",
         "lambda": "λ", "mu": "μ", "pi": "π", "rho": "ρ", "sigma": "σ", "tau": "τ", "phi": "φ", "omega": "ω",
         "Delta": "Δ", "Sigma": "Σ", "Omega": "Ω", "times": "×", "cdot": "·", "leq": "≤", "geq": "≥",
         "approx": "≈", "neq": "≠", "pm": "±", "infty": "∞", "rightarrow": "→", "to": "→", "in": "∈"}


def inline_math(tex):
    """Short inline math as text: $\\mu = 0.01$ → μ = 0.01. Anything bigger belongs in a $$...$$ line."""
    t = re.sub(r"\\([A-Za-z]+)", lambda m: GREEK.get(m.group(1), m.group(1)), tex)
    t = re.sub(r"\^\{?2\}?", "²", t)
    t = re.sub(r"\^\{?3\}?", "³", t)
    return re.sub(r"[{}]", "", t)


def render_equation(tex, out_dir: Path):
    """$$ LaTeX $$ → PNG (matplotlib mathtext: no LaTeX install needed), 32 pt at 300 dpi."""
    import hashlib
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    out_dir.mkdir(parents=True, exist_ok=True)
    png = out_dir / f"eq-{hashlib.sha1(('40' + tex).encode()).hexdigest()[:10]}.png"
    if not png.exists():
        fig = plt.figure()
        fig.text(0, 0, f"${tex}$", fontsize=40, math_fontfamily="stix")
        fig.savefig(png, dpi=300, bbox_inches="tight", pad_inches=0.05, transparent=True)
        plt.close(fig)
    return png


class Cites:
    def __init__(self, bib):
        self.bib, self.order, self.missing = bib, [], set()

    def number(self, keys):
        out = []
        for k in keys:
            if k not in self.bib:
                self.missing.add(k)
                out.append("?")
                continue
            if k not in self.order:
                self.order.append(k)
            out.append(str(self.order.index(k) + 1))
        return "[" + ", ".join(out) + "]"


def runs(text, cites):
    """Markdown inline → [(text, style)] with style in {"", "b", "i", "bi", "code"}."""
    out, pos = [], 0
    for m in INLINE.finditer(text):
        if m.start() > pos:
            out.append((text[pos:m.start()], ""))
        todo, b, i, i2, code, cite, math = m.groups()
        if todo is not None:
            out.append((todo, "todo"))
        elif b is not None:
            out += [(t, "b" + ("i" if s == "i" else "")) for t, s in runs(b, cites)]
        elif i is not None or i2 is not None:
            out += [(t, "i" + ("b" if s == "b" else "")) for t, s in runs(i or i2, cites)]
        elif code is not None:
            out.append((code, "code"))
        elif cite is not None:
            keys = [k.strip().lstrip("@") for k in cite.split(";")]
            out.append((cites.number(keys), ""))
        elif math is not None:
            out.append((inline_math(math), "i"))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], ""))
    return [(t, s) for t, s in out if t]


# ── fonts and measuring ──────────────────────────────────────
class Fonts:
    """Template theme fonts resolved to installed files (or same-width stand-ins)."""

    def __init__(self, prs, problems):
        theme = [r for r in prs.slide_master.part.rels.values() if r.reltype.endswith("/theme")][0].target_part
        xml = theme.blob.decode("utf8", "replace")
        self.major = (re.search(r'<a:majorFont>\s*<a:latin typeface="([^"]*)"', xml) or [None, "Calibri"])[1]
        self.minor = (re.search(r'<a:minorFont>\s*<a:latin typeface="([^"]*)"', xml) or [None, "Calibri"])[1]
        self.files, self._cache, self.problems = {}, {}, problems

    def resolve(self, family):
        family = {"+mj-lt": self.major, "+mn-lt": self.minor}.get(family, family)
        return family

    def file(self, family, style=""):
        family = self.resolve(family)
        key = (family, style)
        if key not in self.files:
            fc_style = {"": "Regular", "b": "Bold", "i": "Italic", "bi": "Bold Italic"}[style]
            r = subprocess.run(["fc-match", "-f", "%{family}|%{file}", f"{family}:style={fc_style}"],
                               capture_output=True, text=True)
            got, path = (r.stdout.split("|", 1) + [""])[:2]
            names = {n.strip().lower() for n in got.split(",")}
            want = family.lower()
            if want not in names and not (METRIC_TWINS.get(want, set()) & names):
                msg = (f"font “{family}” is not installed (fontconfig offers “{got.split(',')[0]}”, whose letter "
                       f"widths differ): text fitting would be wrong. Install {family}"
                       + (f" or {'/'.join(sorted(METRIC_TWINS[want]))}" if want in METRIC_TWINS else "")
                       + " in ~/.local/share/fonts")
                if ("FAIL", msg) not in self.problems:
                    self.problems.append(("FAIL", msg))
            self.files[key] = path
        return self.files[key]

    def font(self, family, style, size_pt):
        if style == "todo":
            style = ""
        fam = CODE_FONT if style == "code" else family
        key = (fam, style, round(size_pt * 4))
        if key not in self._cache:
            self._cache[key] = ImageFont.truetype(self.file(fam, "" if style == "code" else style),
                                                  size=round(size_pt * 4))
        return self._cache[key]

    def width(self, text, family, style, size_pt):
        return self.font(family, "code" if style == "code" else style, size_pt).getlength(text) / 4


def wrap_lines(rs, fonts, family, size, width):
    """Greedy word wrap of runs into `width` points → number of lines."""
    words = []
    for text, style in rs:
        for w in re.split(r"(\s+)", text):
            if w:
                words.append((w, style))
    lines, x = 1, 0.0
    for w, style in words:
        if w.isspace():
            x += fonts.width(" ", family, "", size)
            continue
        ww = fonts.width(w, family, style, size)
        if x > 0 and x + ww > width:
            lines += 1
            x = 0.0
            if ww > width:                       # one long word: it wraps by itself
                lines += int(ww // width)
                x = ww % width
                continue
        x += ww
    return lines


# ── template geometry ────────────────────────────────────────
def ln_inset(bodypr, name, default):
    v = bodypr.get(name) if bodypr is not None else None
    return int(v) / EMU_PT if v is not None else default


class Template:
    """What the deck builder needs from a .pptx: one layout per slide type and the text styles."""

    TITLE_T, CENTER_T, SUB_T, BODY_T, OBJ_T, PIC_T = 1, 3, 4, 2, 7, 18

    def __init__(self, path: Path, problems):
        self.prs = Presentation(str(path))
        for sid in list(self.prs.slides._sldIdLst):         # drop sample slides a template may carry
            self.prs.part.drop_rel(sid.rId)
            self.prs.slides._sldIdLst.remove(sid)
        self.W, self.H = self.prs.slide_width, self.prs.slide_height
        mapping = path.with_suffix(".yml")           # written by template-extract, checked by a person
        self.mapping = ((yaml.safe_load(mapping.read_text()) or {}).get("layouts") or {}) if mapping.exists() else {}
        self.fonts = Fonts(self.prs, problems)
        self.levels = self._levels()
        self.layout = {k: self._find(k) for k in ("title", "section", "content", "two", "title_only", "blank")}
        missing = [k for k, v in self.layout.items() if v is None and k != "two"]
        if missing:
            problems.append(("WARN", f"template has no layout for {', '.join(missing)}: using the closest one"))
        if self.layout["section"] is None:          # a title slide reads as a divider; a bullet slide doesn't
            self.layout["section"] = self.layout["title"]
        self.layout = {k: v or self.layout["content"] or self.prs.slide_layouts[0] for k, v in self.layout.items()}
        self.palette = self._palette()
        self.content_box = self.clear(self.layout["content"], self._content_box())
        self.has_logo = any(sh.shape_type == 13 for sh in self.decorations(self.layout["title"]))

    def decorations(self, layout):
        """Pictures/shapes the master and layout draw on every slide (logos, banners), not placeholders."""
        shown = layout._element.get("showMasterSp", "1") not in ("0", "false")
        shapes = list(layout.shapes) + (list(self.prs.slide_master.shapes) if shown else [])
        return [sh for sh in shapes if not sh.is_placeholder and sh.width and sh.height
                and not sh.name.startswith("Kit Bg")                      # kit theme backdrops: text sits on them
                and sh.width * sh.height < self.W * self.H * 0.5]          # skip full-slide backgrounds

    def clear(self, layout, box, gap=91440):
        """Shrink a box so text never runs under a logo or banner from the master/layout."""
        l, t, w, h = (int(v) for v in box)
        for sh in self.decorations(layout):
            ol, ot, orr, ob = sh.left, sh.top, sh.left + sh.width, sh.top + sh.height
            if ol >= l + w or orr <= l or ot >= t + h or ob <= t:
                continue                                      # no overlap
            if sh.width > self.W * 0.6:                       # a banner: keep above/below it
                if ot > t + h / 2:
                    h = ot - gap - t
                else:
                    h, t = h - (ob + gap - t), ob + gap
            elif ol > l + w / 2:                              # logo on the right
                w = ol - gap - l
            else:                                             # logo on the left
                w, l = w - (orr + gap - l), orr + gap
        return (l, t, max(w, 914400), max(h, 457200))

    def _palette(self):
        """Theme colours by name (dk1, lt1, dk2, lt2, accent1…6): the template's own colours."""
        theme = [r for r in self.prs.slide_master.part.rels.values() if r.reltype.endswith("/theme")][0].target_part
        xml = theme.blob.decode("utf8", "replace")
        out = {}
        for name in ("dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"):
            m = re.search(rf"<a:{name}>.*?(?:srgbClr val|lastClr)=\"([0-9A-Fa-f]{{6}})\"", xml, flags=re.S)
            out[name] = m.group(1).upper() if m else {"dk1": "000000", "lt1": "FFFFFF"}.get(name, "808080")
        return out

    def background_of(self, container):
        """Solid background colour of a slide/layout/master (None: not solid or not set)."""
        bg = container._element.find(".//" + qn("p:bg"))
        if bg is None:
            return None
        c = bg.find(".//" + qn("a:srgbClr"))
        if c is not None:
            return c.get("val").upper()
        c = bg.find(".//" + qn("a:schemeClr"))
        if c is not None:
            v = {"bg1": "lt1", "bg2": "lt2", "tx1": "dk1", "tx2": "dk2"}.get(c.get("val"), c.get("val"))
            return self.palette.get(v)
        return None

    def is_dark(self, layout):
        c = self.background_of(layout) or self.background_of(self.prs.slide_master) or "FFFFFF"
        import deck_theme
        return deck_theme.luminance(c) < 0.25

    def _types(self, lay):
        return [(ph.placeholder_format.type and int(ph.placeholder_format.type)) for ph in lay.placeholders
                if ph.placeholder_format.type not in (None,) and int(ph.placeholder_format.type) not in (13, 15, 16)]

    def _find(self, kind):
        names = {"title": ["Title Slide"], "section": ["Section Header", "Section Title"],
                 "content": ["Title and Content", "Title, Content"], "two": ["Two Content"],
                 "title_only": ["Title Only"], "blank": ["Blank"]}[kind]
        lays = list(self.prs.slide_layouts)
        if kind in self.mapping:
            want = self.mapping[kind]
            return next((lay for lay in lays if lay.name == want), None) if want else None
        for lay in lays:
            if lay.name in names:
                return lay
        for lay in lays:
            t = self._types(lay)
            body = [x for x in t if x in (self.BODY_T, self.OBJ_T)]
            if kind == "title" and self.CENTER_T in t:
                return lay
            if kind == "content" and self.TITLE_T in t and len(body) == 1:
                return lay
            if kind == "two" and self.TITLE_T in t and len(body) == 2:
                return lay
            if kind == "title_only" and t == [self.TITLE_T]:
                return lay
            if kind == "blank" and not t:
                return lay
        return None

    def _levels(self):
        """Body text levels from the master: size, left margin, indent, space before, line spacing."""
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
              "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
        body = self.prs.slide_master._element.find(".//p:txStyles/p:bodyStyle", ns)
        out = {}
        for lvl in range(1, 6):
            ppr = body.find(f"a:lvl{lvl}pPr", ns) if body is not None else None
            get = lambda path, attr, d: (ppr.find(path, ns).get(attr) if ppr is not None  # noqa: E731
                                         and ppr.find(path, ns) is not None else d)
            sz = int(get("a:defRPr", "sz", 2400 - 200 * lvl)) / 100
            spc_pts = get("a:spcBef/a:spcPts", "val", None)
            spc_pct = get("a:spcBef/a:spcPct", "val", None)
            ln_pct = get("a:lnSpc/a:spcPct", "val", "100000")
            out[lvl - 1] = {
                "size": sz,
                "marL": int(ppr.get("marL", 0) if ppr is not None else 0) / EMU_PT,
                "spc": (int(spc_pts) / 100) if spc_pts else ("pct", int(spc_pct) / 100000) if spc_pct else 0,
                "ln": int(ln_pct) / 100000,
            }
        return out

    def _content_box(self):
        lay = self.layout["content"]
        for ph in lay.placeholders:
            if ph.placeholder_format.idx == 1:
                return (ph.left, ph.top, ph.width, ph.height)
        t = self.title_ph(self.layout["title_only"])
        top = (t.top + t.height + Emu(91440)) if t is not None else int(self.H * 0.2)
        return (int(self.W * 0.05), top, int(self.W * 0.9), int(self.H * 0.92) - top)

    @staticmethod
    def title_ph(container):
        for ph in container.placeholders:
            if ph.placeholder_format.idx == 0:
                return ph
        return None


# ── writing text with measured sizes ─────────────────────────
def set_runs(par, rs, size, colour=None):
    for text, style in rs:
        r = par.add_run()
        r.text = text
        r.font.size = Pt(size)
        if style == "todo":                       # gaps stay visible: yellow, like in the report
            rpr = r._r.get_or_add_rPr()
            hl = etree.SubElement(rpr, qn("a:highlight"))
            etree.SubElement(hl, qn("a:srgbClr"), val="FFFF00")
            continue
        if "b" in style and style != "code":
            r.font.bold = True
        if "i" in style:
            r.font.italic = True
        if style == "code":
            r.font.name = CODE_FONT
        if colour:
            r.font.color.rgb = RGBColor.from_string(colour)


def no_bullet(par):
    ppr = par._p.get_or_add_pPr()
    ppr.set("marL", "0")
    ppr.set("indent", "0")
    for tag in ("a:buNone", "a:buChar", "a:buAutoNum"):
        for e in ppr.findall(qn(tag)):
            ppr.remove(e)
    etree.SubElement(ppr, qn("a:buNone"))


def text_height(paras, tpl, family, scale, width_pt, levels=None):
    """paras: [(level, runs, bullet)] → needed height in points at `scale` × level sizes."""
    levels = levels or tpl.levels
    h = 0.0
    for k, (lvl, rs, bullet) in enumerate(paras):
        L = levels.get(lvl, levels[max(levels)])
        size = L["size"] * scale
        indent = L["marL"] if bullet else 0
        lines = wrap_lines(rs, tpl.fonts, family, size, max(width_pt - indent, 20))
        spc = L["spc"]
        spc = size * spc[1] if isinstance(spc, tuple) else spc
        h += (spc if k else 0) + lines * size * 1.2 * L["ln"]
    return h


def fill_text(shape, paras, tpl, problems, where, family="+mn-lt", min_size=MIN_BODY, levels=None,
              colour=None, anchor=None, align=None, max_size=None):
    """Put paragraphs into a text shape at the largest scale (down to min_size for level 0) that fits."""
    levels = levels or tpl.levels
    tf = shape.text_frame
    tf.word_wrap = True
    tf.auto_size = MSO_AUTO_SIZE.NONE
    if anchor:
        tf.vertical_anchor = anchor
    bp = tf._txBody.find(qn("a:bodyPr"))
    lr = ln_inset(bp, "lIns", 7.2) + ln_inset(bp, "rIns", 7.2)
    tb = ln_inset(bp, "tIns", 3.6) + ln_inset(bp, "bIns", 3.6)
    width, height = shape.width / EMU_PT - lr, (shape.height / EMU_PT - tb) * SAFETY
    base = levels[0]["size"]
    scale, fits = min(1.0, (max_size or base) / base), False
    while True:
        if text_height(paras, tpl, family, scale, width, levels) <= height:
            fits = True
            break
        nxt = (base * scale - 2) / base
        if base * nxt < min_size - 0.01:
            break
        scale = nxt
    if not fits:
        problems.append(("FAIL", f"{where}: text does not fit even at {round(base * scale)} pt "
                                 f"(needs {text_height(paras, tpl, family, scale, width, levels):.0f} pt of "
                                 f"{height:.0f} pt): cut words or split the slide"))
    tf.clear()
    for k, (lvl, rs, bullet) in enumerate(paras):
        par = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
        par.level = min(lvl, 4)
        if not bullet:
            no_bullet(par)
        if align:
            par.alignment = align
        size = round(levels.get(lvl, levels[max(levels)])["size"] * scale)
        set_runs(par, rs, size, colour)
    return round(base * scale)


def title_text(ph, text, tpl, cites, problems, where, min_size=MIN_TITLE, size=None, align=None, colour=None):
    size = size or font_size_of(ph, tpl, default=32)
    lv = {0: {"size": size, "marL": 0, "spc": 0, "ln": 1.0}}
    got = fill_text(ph, [(0, runs(text, cites), False)], tpl, problems, where + " title", family="+mj-lt",
                    min_size=min(min_size, size), levels=lv, align=align, colour=colour)
    if got < min(size, 28):            # big template titles may shrink a little; small ones are hard to read
        problems.append(("WARN", f"{where}: title shrunk from {size:.0f} to {got} pt to fit; shorten it"))
    return got


def font_size_of(ph, tpl, default):
    """Size a placeholder inherits: its own lstStyle, then the layout's, then the master title style."""
    sz = None
    try:
        lay_ph = ph.part.slide.slide_layout.placeholders.get(idx=ph.placeholder_format.idx)
    except Exception:
        lay_ph = None
    for cand in (ph, lay_ph):
        if cand is None:
            continue
        d = cand._element.find(".//" + qn("a:lstStyle") + "/" + qn("a:lvl1pPr") + "/" + qn("a:defRPr"))
        if d is not None and d.get("sz"):
            sz = int(d.get("sz")) / 100
            break
    if sz is None:
        ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
              "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
        d = tpl.prs.slide_master._element.find(".//p:txStyles/p:titleStyle/a:lvl1pPr/a:defRPr", ns)
        sz = int(d.get("sz")) / 100 if d is not None and d.get("sz") else default
    return sz


# ── pictures (fit, never crop) ───────────────────────────────
def resolve_image(path, sdir, cfg):
    roots = [sdir, sdir / "figures"]
    if cfg["_report"]:
        roots += [cfg["_report"], cfg["_report"] / "figures"]
    p = Path(path)
    cands = [p] if p.suffix and p.suffix.lower() != ".svg" else \
        [p.with_suffix(s) for s in (".png", ".jpg", ".jpeg")]
    for root in roots:
        for c in cands:
            f = c if c.is_absolute() else root / c
            if f.exists():
                return f
    return None


def place_picture(slide, img, box, problems, where, valign="middle", natural_dpi=None):
    """Scale to fit inside box keeping the aspect ratio (never crop); centre it."""
    l, t, w, h = box
    with Image.open(img) as im:
        iw, ih = im.size
    s = min(w / iw, h / ih)
    if natural_dpi:                                   # e.g. equations: no bigger than drawn
        s = min(s, 914400 / natural_dpi)
    pw, ph = int(iw * s), int(ih * s)
    left = l + (w - pw) // 2
    top = t + ((h - ph) // 2 if valign == "middle" else 0)
    pic = slide.shapes.add_picture(str(img), left, top, pw, ph)
    placed_in = pw / 914400
    import deck_diagrams
    small = deck_diagrams.smallest_text_pt(img, pw / EMU_PT)
    if small is not None and small < 14:
        problems.append(("WARN", f"{where}: text in {img.name} is only {small:.0f} pt on the slide; draw a slide "
                                 f"version in slides/diagrams/ (fewer nodes, left-to-right)"))
    if iw / placed_in < 110:
        problems.append(("WARN", f"{where}: {img.name} is only {iw / placed_in:.0f} px per inch on the slide "
                                 f"(blurry when projected); export it larger"))
    return pic


def remove_shape(shape):
    shape._element.getparent().remove(shape._element)


# ── footer and slide number ──────────────────────────────────
def copy_footer(slide, layout, footer_text):
    """python-pptx doesn't carry footer/slide-number placeholders onto slides; copy them from the layout."""
    tree = slide.shapes._spTree
    next_id = max([int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))] + [1]) + 1
    for ph in layout.placeholders:
        t = ph.placeholder_format.type
        if t is None or int(t) not in (13, 15):          # slide number, footer
            continue
        if int(t) == 15 and not footer_text:
            continue
        el = copy.deepcopy(ph._element)
        el.find(".//" + qn("p:cNvPr")).set("id", str(next_id))
        next_id += 1
        tree.append(el)
        if int(t) == 15:
            body = el.find(qn("p:txBody"))
            for p in body.findall(qn("a:p")):
                body.remove(p)
            p = etree.SubElement(body, qn("a:p"))
            r = etree.SubElement(p, qn("a:r"))
            rpr = etree.SubElement(r, qn("a:rPr"), lang="en-US")
            etree.SubElement(rpr, qn("a:latin"), typeface="+mn-lt")
            etree.SubElement(r, qn("a:t")).text = footer_text


# ── slide builders ───────────────────────────────────────────
IN = 914400
LOOKS = ("dark", "accent", "light")


def set_alpha(shape, alpha):
    c = shape._element.spPr.find(".//" + qn("a:srgbClr"))
    if c is not None:
        etree.SubElement(c, qn("a:alpha"), val=str(int(alpha * 100000)))


def to_back(shape, behind=0):
    tree = shape._element.getparent()
    tree.remove(shape._element)
    tree.insert(2 + behind, shape._element)


def recolour(slide, colour):
    """Text without its own colour → `colour` (a slide given a dark/accent/light background)."""
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        for p in sh.text_frame.paragraphs:
            for r in list(p._p.findall(qn("a:r"))) + list(p._p.findall(qn("a:fld"))):
                rpr = r.find(qn("a:rPr"))
                if rpr is None:
                    rpr = etree.Element(qn("a:rPr"))
                    r.insert(0, rpr)
                if rpr.find(qn("a:solidFill")) is not None:
                    continue
                fill = etree.Element(qn("a:solidFill"))
                todo = rpr.find(qn("a:highlight")) is not None
                etree.SubElement(fill, qn("a:srgbClr"), val="000000" if todo else colour)
                rpr.insert(0, fill)
            if sh.is_placeholder and p.text.strip():           # bullets in the theme colour may vanish too
                ppr = p._p.get_or_add_pPr()
                for e in ppr.findall(qn("a:buClr")):
                    ppr.remove(e)
                bu = etree.Element(qn("a:buClr"))
                etree.SubElement(bu, qn("a:srgbClr"), val=colour)
                ppr.insert(0, bu)


def split_head(text):
    """'**Privacy**: data stays local' / 'Privacy: data stays local' → ('Privacy', 'data stays local')."""
    m = re.match(r"\*\*(.+?)\*\*\s*[:—–-]?\s*(.*)$", text)
    if m:
        return m.group(1), m.group(2)
    m = re.match(r"([^:`$\[]{1,40}):\s+(.+)$", text)
    if m and len(m.group(1).split()) <= 5:
        return m.group(1), m.group(2)
    return None, text


class Builder:
    def __init__(self, sdir, cfg, tpl, cites, problems):
        self.sdir, self.cfg, self.tpl, self.cites, self.problems = sdir, cfg, tpl, cites, problems
        self.prs = tpl.prs
        self.footer = cfg.get("footer", cfg.get("title") if len(cfg.get("title", "")) <= 70 else "")
        self.cur = None             # the slides.md slide being built: its {.dark} {bg=…} apply
        self.dark = False           # the slide being built has a dark background
        self.forced = None          # text colour a per-slide background forces
        self.slide_bg = "FFFFFF"
        self.flip = False           # picture side alternates on picture + text slides
        self.sections, self.progress = [], False    # `# Section` names: agenda, and "you are here" on dividers
        self.variety = str(cfg.get("variety", "auto")).lower() not in ("false", "off", "none", "no")
        logo = cfg["me"].get("logo")
        self.logo = Path(logo) if logo and Path(logo).exists() else None

    def new(self, kind, look=None, cover=False):
        lay = self.tpl.layout[kind]
        slide = self.prs.slides.add_slide(lay)
        for ph in slide.placeholders:                 # keep title/body clear of the template's logo/banner
            t = ph.placeholder_format.type
            if t is not None and int(t) in (1, 2, 3, 4, 7, 18):
                box = (ph.left, ph.top, ph.width, ph.height)
                new = self.tpl.clear(lay, box)
                if new != tuple(int(v) for v in box):
                    ph.left, ph.top, ph.width, ph.height = new
        self.slide_bg = (self.tpl.background_of(lay) or self.tpl.background_of(self.prs.slide_master)
                         or self.tpl.palette["lt1"])
        self.dark = self.tpl.is_dark(lay)
        self.forced = None
        if kind != "title" and not cover:
            self.apply_look(slide, look)
            if kind != "section":
                self.small_logo(slide)
        return slide

    def apply_look(self, slide, default=None):
        """{.dark} {.accent} {.light} {bg=photo.jpg} {color="#hex"}: this slide's own background."""
        import deck_theme
        sl = self.cur or {}
        cls, attrs = set(sl.get("classes", [])), sl.get("attrs", {})
        P = self.tpl.palette
        look = next((c for c in LOOKS if c in cls), None)
        colour = attrs.get("color") or attrs.get("colour")
        if colour:
            look = "colour"
        if attrs.get("bg"):
            look = "image"
        look = look or default
        if not look:
            return
        slide._element.set("showMasterSp", "0")         # the theme's bars and bands would clash with it
        bg = {"dark": P["dk2"], "accent": P["accent1"], "light": P["lt2"] if self.dark else "FFFFFF",
              "colour": (colour or "").lstrip("#").upper(), "image": "111111"}[look]
        if not re.fullmatch(r"[0-9A-F]{6}", bg):
            self.problems.append(("FAIL", f"slides.md:{sl.get('line')}: color={colour!r} is not a hex colour "
                                          f"like \"#7B2CBF\""))
            bg = P["dk2"]
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor.from_string(bg)
        if look == "image":
            img = resolve_image(attrs["bg"], self.sdir, self.cfg)
            if img is None:
                self.problems.append(("FAIL", f"slides.md:{sl.get('line')}: background {attrs['bg']} not found"))
            else:
                shade = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, self.tpl.W, self.tpl.H)
                self.plain_shape(shade, "000000", alpha=float(attrs.get("dim", 0.5)))
                to_back(shade)
                self.cover_picture(slide, img)
        self.slide_bg = bg
        self.dark = deck_theme.luminance(bg) < 0.3
        self.forced = "FFFFFF" if self.dark else P["dk1"]

    def cover_picture(self, slide, img):
        """A photo filling the whole slide (cropped to fill: only for backgrounds)."""
        W, H = self.tpl.W, self.tpl.H
        with Image.open(img) as im:
            iw, ih = im.size
        pic = slide.shapes.add_picture(str(img), 0, 0, W, H)
        if iw / ih > W / H:
            pic.crop_left = pic.crop_right = (1 - (W / H) / (iw / ih)) / 2
        else:
            pic.crop_top = pic.crop_bottom = (1 - (iw / ih) / (W / H)) / 2
        to_back(pic)

    def small_logo(self, slide):
        """logo_on_slides: true → a small logo top-right on every content slide; the title steps aside."""
        if not (self.cfg.get("logo_on_slides") and self.logo) or self.tpl.has_logo:
            return
        with Image.open(self.logo) as im:
            iw, ih = im.size
        h = int(0.7 * IN)
        w = min(int(h * iw / ih), int(1.6 * IN))
        h = int(w * ih / iw)
        x = self.tpl.W - int(0.45 * IN) - w
        slide.shapes.add_picture(str(self.logo), x, int(0.3 * IN), w, h)
        t = slide.shapes.title
        if t is not None and t.left + t.width > x - int(0.15 * IN):
            # all four at once: setting one on an inherited position zeroes the others
            t.left, t.top, t.width, t.height = t.left, t.top, x - int(0.15 * IN) - t.left, t.height

    def ink(self):
        """Colours for things the builder draws, right for this slide's background."""
        import deck_theme
        P, bg = self.tpl.palette, self.slide_bg
        if self.dark:
            return {"title": "FFFFFF", "text": "F2F2F2", "muted": "C9CFD6", "card": deck_theme.mix(bg, "FFFFFF", 0.1),
                    "number": deck_theme.mix(P["accent2"], "FFFFFF", 0.15), "line": deck_theme.mix(bg, "FFFFFF", 0.4),
                    "on_accent": "FFFFFF",
                    "accent": P["accent1"] if deck_theme.luminance(P["accent1"]) > 0.25
                    else deck_theme.mix(P["accent1"], "FFFFFF", 0.45),
                    "accents": [P["accent2"], deck_theme.mix(P["accent1"], "FFFFFF", 0.3), P["accent3"], P["accent4"]]}
        card = P["lt2"] if bg == "FFFFFF" else "FFFFFF"
        if deck_theme.luminance(card) < 0.6:                 # a template whose lt2 is dark
            card = deck_theme.mix(P["accent1"], "FFFFFF", 0.9)
        return {"title": P["dk2"], "text": P["dk1"], "muted": "595959", "card": card, "number": P["accent1"],
                "line": "BFBFBF", "on_accent": "FFFFFF", "accent": P["accent1"],
                "accents": [P["accent1"], P["accent2"], P["accent3"], P["accent4"]]}

    def plain_shape(self, sh, fill, alpha=None, name=None):
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(fill)
        sh.line.fill.background()
        st = sh._element.find(qn("p:style"))
        if st is not None:                       # the Office shape style adds an outline and a shadow
            sh._element.remove(st)
        if alpha is not None:
            set_alpha(sh, alpha)
        if name:
            sh.name = name
        return sh

    def shape(self, s, kind, box, fill, alpha=None, name=None):
        return self.plain_shape(s.shapes.add_shape(kind, *(int(v) for v in box)), fill, alpha, name)

    def text(self, s, box, paras, size, n, sl, colour=None, family="+mn-lt", min_size=14, align=None,
             anchor=MSO_ANCHOR.TOP, shape=None, sizes=None, what=""):
        """Plain text at a measured size: paras [(level, runs)], level 1 = smaller lines under level 0."""
        tb = shape or s.shapes.add_textbox(*(int(v) for v in box))
        lv = {lvl: {"size": sz, "marL": 0, "spc": sz * 0.3 if lvl else 0, "ln": 1.0}
              for lvl, sz in (sizes or {0: size}).items()}
        got = fill_text(tb, [(lvl, rs, False) for lvl, rs in paras], self.tpl, self.problems,
                        self.where(n, sl) + what, family=family, min_size=min_size, levels=lv, colour=colour,
                        align=align, anchor=anchor)
        return tb, got

    def where(self, n, sl):
        return f"slide {n} “{sl['title'][:40]}” (slides.md:{sl['line']})"

    def finish(self, slide, layout_kind, notes, footer=True):
        if footer:
            copy_footer(slide, self.tpl.layout[layout_kind], self.footer if footer is True else "")
        for ph in list(slide.placeholders):          # unused placeholders show "Click to add text" in editors
            if ph.has_text_frame and not ph.text_frame.text.strip() and ph.placeholder_format.idx not in (13,):
                if not getattr(ph, "_kit_keep", False):
                    remove_shape(ph)
        if self.forced:
            recolour(slide, self.forced)
        if notes:
            slide.notes_slide.notes_text_frame.text = "\n".join(notes).strip()

    # title
    def title_slide(self):
        cover = str(self.cfg.get("cover") or "theme")
        if cover.lower().endswith(".pptx"):
            return self.cover_from_file(cover)
        if cover == "college":
            return self.college_cover()
        cfg, me = self.cfg, self.cfg["me"]
        s = self.new("title")
        where = "title slide"
        title_text(s.shapes.title, cfg["title"], self.tpl, self.cites, self.problems, where, min_size=28)
        lines = []
        if cfg.get("subtitle"):
            lines.append((0, [(cfg["subtitle"], "b")], False))
        team = self.team()
        if team:
            lines.append((0, [("Presented by ", "")] + [x for k, (nm, roll) in enumerate(team)
                                                         for x in ([(", ", "")] if k else []) + [(nm, "b")]
                                                         + ([(f" ({roll})", "")] if roll else [])], False))
        g = me.get("guide") or {}
        if g.get("name"):
            lines.append((0, [("Guide: ", ""), (g["name"], "b")] + ([(f", {g['title']}", "")] if g.get("title") else []), False))
        org = ", ".join(x for x in (me.get("department"), me.get("college")) if x)
        if org:
            lines.append((0, [(org, "")], False))
        lines.append((0, [(str(cfg["date"]), "")], False))
        sub = next((ph for ph in s.placeholders if ph.placeholder_format.idx == 1), None)
        if sub is not None and lines:
            # the details may need more room than the template's subtitle box: extend it to the free space below
            bottom = self.tpl.clear(self.tpl.layout["title"], (sub.left, sub.top, sub.width,
                                                             int(self.tpl.H * 0.95) - sub.top))
            if bottom[3] > sub.height:
                sub.left, sub.top, sub.width, sub.height = bottom
            size = min(font_size_of(sub, self.tpl, 20), 24)
            lv = {0: {"size": size, "marL": 0, "spc": size * 0.35, "ln": 1.0}}
            fill_text(sub, lines, self.tpl, self.problems, where + " details", min_size=14, levels=lv)
        if self.logo and cfg.get("logo", True) and not self.tpl.has_logo:
            W, H = self.tpl.W, self.tpl.H
            box = (int(W - W * 0.05 - 1.4 * IN), int(H * 0.06), int(1.4 * IN), int(1.4 * IN))
            place_picture(s, self.logo, box, self.problems, where, valign="top")
        self.finish(s, "title", [], footer=False)

    def team(self):
        """[(name, roll)]: slides.yml `team:` (names or {name, roll}), else me.yml."""
        me = self.cfg["me"]
        team = self.cfg.get("team")
        if team:
            return [(t, "") if isinstance(t, str) else (t.get("name", ""), str(t.get("roll", "") or ""))
                    for t in team]
        return [(me["name"], str(me.get("roll") or ""))] if me.get("name") else []

    def cover_base(self):
        """A plain slide in the background of the theme's content slides (white, tinted or dark)."""
        s = self.new("title_only", cover=True)
        s._element.set("showMasterSp", "0")
        lay = self.tpl.layout["content"]
        bg = self.tpl.background_of(lay) or self.tpl.background_of(self.prs.slide_master) or "FFFFFF"
        s.background.fill.solid()
        s.background.fill.fore_color.rgb = RGBColor.from_string(bg)
        self.slide_bg, self.dark = bg, self.tpl.is_dark(lay)
        return s

    def logo_plate_behind(self, s, pic):
        """A logo on a dark slide sits on a white plate (most logos have a white background)."""
        if self.dark:
            plate = self.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, (pic.left - 0.12 * IN, pic.top - 0.1 * IN,
                                                                pic.width + 0.24 * IN, pic.height + 0.2 * IN), "FFFFFF")
            plate.adjustments[0] = 0.12
            el = plate._element                       # behind the logo
            el.getparent().remove(el)
            pic._element.addprevious(el)

    def college_cover(self):
        """cover: college → the usual college first slide: logo and college name on top, title in the
        middle, "Presented by" and "Under the guidance of" at the bottom. White, in the theme's colours."""
        cfg, me, W, H = self.cfg, self.cfg["me"], self.tpl.W, self.tpl.H
        s = self.cover_base()
        ink, P, where = self.ink(), self.tpl.palette, "title slide"
        sl = {"title": "title slide", "line": 0}
        y = int(0.3 * IN)
        if self.logo:
            pic = place_picture(s, self.logo, (int(W / 2 - 0.9 * IN), y, int(1.8 * IN), int(1.25 * IN)),
                                self.problems, where, valign="top")
            self.logo_plate_behind(s, pic)
            y += int(1.35 * IN)
        else:
            self.problems.append(("WARN", "cover: college has no logo: set logo: in me.yml (path to the image)"))
        mx = int(0.6 * IN)
        if me.get("college"):
            self.text(s, (mx, y, W - 2 * mx, int(0.55 * IN)), [(0, [(me["college"], "b")])], 24, 0, sl,
                      colour=ink["title"], family="+mj-lt", align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            y += int(0.55 * IN)
        if me.get("department"):
            self.text(s, (mx, y, W - 2 * mx, int(0.42 * IN)), [(0, [(me["department"], "")])], 18, 0, sl,
                      colour=ink["muted"], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            y += int(0.5 * IN)
        self.shape(s, MSO_SHAPE.RECTANGLE, (W * 0.3, y, W * 0.4, int(0.035 * IN)), ink["accent"])
        y += int(0.2 * IN)
        t = s.shapes.title
        t.left, t.top, t.width, t.height = int(W * 0.07), y, int(W * 0.86), int(1.45 * IN)
        title_text(t, cfg["title"], self.tpl, self.cites, self.problems, where, min_size=26, size=36,
                   align=PP_ALIGN.CENTER, colour=ink["accent"])
        t.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        y += int(1.45 * IN)
        if cfg.get("subtitle"):
            self.text(s, (mx, y, W - 2 * mx, int(0.45 * IN)), [(0, [(cfg["subtitle"], "i")])], 20, 0, sl,
                      colour=ink["text"], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            y += int(0.5 * IN)
        top = max(y + int(0.15 * IN), H - int(2.15 * IN))
        col_w, bottom_h = int((W - 2 * mx) * 0.46), H - int(0.6 * IN) - top
        team = self.team()
        left = [(0, [("Presented by", "b")])] + [(1, [(nm, "b")] + ([(f"  ({roll})", "")] if roll else []))
                                                  for nm, roll in team]
        g = me.get("guide") or {}
        right = ([(0, [("Under the guidance of", "b")]), (1, [(g["name"], "b")])]
                 + ([(1, [(g["title"], "")])] if g.get("title") else [])) if g.get("name") else []
        if team:
            self.text(s, (mx, top, col_w, bottom_h), left, 18, 0, sl, colour=ink["text"], min_size=12,
                      sizes={0: 18, 1: 18}, anchor=MSO_ANCHOR.TOP)
        if right:
            self.text(s, (W - mx - col_w, top, col_w, bottom_h), right, 18, 0, sl, colour=ink["text"], min_size=12,
                      sizes={0: 18, 1: 18}, align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.TOP)
        self.text(s, (mx, H - int(0.55 * IN), W - 2 * mx, int(0.4 * IN)), [(0, [(str(cfg["date"]), "")])], 12, 0, sl,
                  colour=ink["muted"], min_size=10, align=PP_ALIGN.CENTER)
        for shp in s.shapes:                       # the heading lines are in the accent colour
            if shp.has_text_frame and shp.text_frame.text.startswith(("Presented by", "Under the guidance")):
                for r in shp.text_frame.paragraphs[0].runs:
                    r.font.color.rgb = RGBColor.from_string(ink["accent"])
        self.finish(s, "title_only", [], footer=False)

    def cover_from_file(self, name):
        """cover: first-page.pptx → its first slide, copied as it is; {{title}}, {{name}}, {{guide}} … filled in."""
        path = next((p for p in (self.sdir / name, Path(name).expanduser()) if p.exists()), None)
        if path is None:
            self.problems.append(("FAIL", f"cover: {name} not found in slides/"))
            return
        src = Presentation(str(path))
        if not len(src.slides._sldIdLst):
            self.problems.append(("FAIL", f"cover: {name} has no slides"))
            return
        ss = src.slides[0]
        s = self.new("blank", cover=True)
        s._element.set("showMasterSp", "0")
        if abs(src.slide_width / self.tpl.W - 1) > 0.01 or abs(src.slide_height / self.tpl.H - 1) > 0.01:
            self.problems.append(("WARN", f"cover: {name} is {src.slide_width / IN:.2f} × {src.slide_height / IN:.2f} in, "
                                          f"the deck {self.tpl.W / IN:.2f} × {self.tpl.H / IN:.2f} in: it will not fill "
                                          f"the slide exactly"))
        sx, sy = self.tpl.W / src.slide_width, self.tpl.H / src.slide_height
        # what the source slide shows: its layout's and master's pictures/shapes (a logo often lives there) + its own
        shown = []
        if ss.slide_layout._element.get("showMasterSp", "1") not in ("0", "false"):
            shown += [(sh, src.slide_master.part) for sh in src.slide_master.shapes if not sh.is_placeholder]
        shown += [(sh, ss.slide_layout.part) for sh in ss.slide_layout.shapes if not sh.is_placeholder]
        shown += [(sh, ss.part) for sh in ss.shapes]
        owner = next((c for c in (ss, ss.slide_layout, src.slide_master)
                      if c._element.find(qn("p:cSld")).find(qn("p:bg")) is not None), None)
        if owner is not None:
            bg = copy.deepcopy(owner._element.find(qn("p:cSld")).find(qn("p:bg")))
            self.relink(bg, owner.part, s.part)
            s._element.find(qn("p:cSld")).insert(0, bg)
        tree = s.shapes._spTree
        next_id = max([int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))] + [1]) + 1
        for sh, part in shown:
            if sh.is_placeholder:                    # make its inherited position explicit, then a plain shape
                sh.left, sh.top, sh.width, sh.height = sh.left, sh.top, sh.width, sh.height
            el = copy.deepcopy(sh._element)
            for ph in el.iter(qn("p:ph")):
                ph.getparent().remove(ph)
            self.relink(el, part, s.part)
            for e in el.iter(qn("p:cNvPr")):
                e.set("id", str(next_id))
                next_id += 1
            if sx != 1 or sy != 1:
                for off in el.iter(qn("a:off")):
                    off.set("x", str(int(int(off.get("x")) * sx)))
                    off.set("y", str(int(int(off.get("y")) * sy)))
                for ext in el.iter(qn("a:ext")):
                    if ext.get("cx") is not None:
                        ext.set("cx", str(int(int(ext.get("cx")) * sx)))
                        ext.set("cy", str(int(int(ext.get("cy")) * sy)))
            tree.append(el)
        me, g = self.cfg["me"], self.cfg["me"].get("guide") or {}
        team = self.team()
        fields = {"title": self.cfg["title"], "subtitle": self.cfg.get("subtitle", ""),
                  "name": ", ".join(nm for nm, _ in team), "roll": ", ".join(r for _, r in team if r),
                  "team": ", ".join(f"{nm} ({r})" if r else nm for nm, r in team),
                  "guide": g.get("name", ""), "guide_title": g.get("title", ""),
                  "department": me.get("department", ""), "college": me.get("college", ""),
                  "date": str(self.cfg["date"])}
        used = set()
        for shp in s.shapes:
            if not shp.has_text_frame:
                continue
            for p in shp.text_frame.paragraphs:
                text = "".join(r.text for r in p.runs)
                if "{{" not in text:
                    continue
                new = re.sub(r"\{\{\s*(\w+)\s*\}\}", lambda m: (used.add(m.group(1)), fields.get(m.group(1), m.group(0)))[1],
                             text)
                p.runs[0].text = new
                for r in p.runs[1:]:
                    r.text = ""
        if "title" not in used:
            self.problems.append(("WARN", f"cover: {name} has no {{{{title}}}} field: type {{{{title}}}} where the "
                                          f"project title goes (also {{{{name}}}}, {{{{roll}}}}, {{{{guide}}}}, "
                                          f"{{{{date}}}} …)"))
        self.finish(s, "blank", [], footer=False)

    @staticmethod
    def relink(el, src_part, dst_part):
        """Relationship ids inside copied XML (pictures, links) → the same targets from the new slide."""
        for e in el.iter():
            for attr, rid in list(e.attrib.items()):
                if not attr.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"):
                    continue
                rel = src_part.rels.get(rid)
                if rel is None:
                    continue
                if rel.is_external:
                    e.set(attr, dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True))
                elif rel.reltype.endswith("/image"):     # a new image part in this package (names would clash)
                    import io
                    e.set(attr, dst_part.get_or_add_image_part(io.BytesIO(rel.target_part.blob))[1])


    def section(self, n, sl, label=None, current=None):
        s = self.new("section")
        if current is not None and current < len(self.sections):   # "you are here" list beside the title
            import deck_patterns
            deck_patterns.progress(self, s, self.sections, current)
            body = next((ph for ph in s.placeholders if ph.placeholder_format.idx == 1), None)
            if body is not None:
                body.left, body.top, body.width, body.height = body.left, body.top, s.shapes.title.width, body.height
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        body = next((ph for ph in s.placeholders if ph.placeholder_format.idx == 1), None)
        if body is not None and label:
            size = font_size_of(body, self.tpl, 20)
            lv = {0: {"size": size, "marL": 0, "spc": 0, "ln": 1.0}}
            fill_text(body, [(0, [(label, "")], False)], self.tpl, self.problems, self.where(n, sl),
                      min_size=12, levels=lv)
        self.finish(s, "section", sl["notes"])
        return s

    def paras_of(self, blocks):
        out = []
        for b in blocks:
            if b["kind"] == "bullets":
                out += [(lvl, runs(t, self.cites), True) for lvl, t in b["items"]]
            elif b["kind"] == "para":
                out.append((0, runs(b["text"], self.cites), False))
        return out

    def check_density(self, n, sl, paras):
        bullets = [p for p in paras if p[2]]
        words = sum(len("".join(t for t, _ in rs).split()) for _, rs, _ in paras)
        if len(bullets) > 6:
            self.problems.append(("WARN", f"{self.where(n, sl)}: {len(bullets)} bullets (aim for ≤ 6)"))
        if words > 40:
            self.problems.append(("WARN", f"{self.where(n, sl)}: {words} words (aim for ≤ 40; "
                                          f"put the rest in ::: notes)"))

    def bullets(self, n, sl):
        paras = self.paras_of(sl["blocks"])
        self.check_density(n, sl, paras)
        chunks = self.split_to_fit(paras)
        if len(chunks) > 1:
            self.problems.append(("WARN", f"{self.where(n, sl)}: too much text for one slide even at "
                                          f"{MIN_BODY} pt: split into {len(chunks)} slides (better: cut it)"))
        bodies = []
        for k, chunk in enumerate(chunks):
            s = self.new("content")
            title = sl["title"] + (" (cont.)" if k else "")
            title_text(s.shapes.title, title, self.tpl, self.cites, self.problems, self.where(n, sl))
            body = next(ph for ph in s.placeholders if ph.placeholder_format.idx == 1)
            if chunk:
                bodies.append((body, chunk, fill_text(body, chunk, self.tpl, self.problems, self.where(n, sl))))
            self.finish(s, "content", sl["notes"] if k == 0 else [])
        smallest = min((size for _, _, size in bodies), default=None)
        for body, chunk, size in bodies:      # continuation slides use one text size, not one each
            if size != smallest:
                fill_text(body, chunk, self.tpl, self.problems, self.where(n, sl), max_size=smallest)

    def split_to_fit(self, paras):
        """Paragraph groups (a top-level item and its sub-items) packed into as few slides as fit at MIN_BODY."""
        l, t, w, h = self.tpl.content_box
        width, height = w / EMU_PT - 14.4, (h / EMU_PT - 7.2) * SAFETY
        scale = MIN_BODY / self.tpl.levels[0]["size"]
        fits = lambda ps: text_height(ps, self.tpl, "+mn-lt", scale, width) <= height  # noqa: E731
        if not paras or fits(paras):
            return [paras]
        groups = []
        for p in paras:
            if p[0] == 0 or not groups:
                groups.append([p])
            else:
                groups[-1].append(p)
        chunks, cur = [], []
        for g in groups:
            if cur and not fits(cur + g):
                chunks.append(cur)
                cur = []
            cur = cur + g
        chunks.append(cur)
        return chunks

    def column(self, s, box, blocks, n, sl, ph=None):
        """Fill one column box with text (placeholder if given) or a picture/table."""
        imgs = [b for b in blocks if b["kind"] == "image"]
        tables = [b for b in blocks if b["kind"] == "table"]
        paras = self.paras_of(blocks)
        if imgs and not paras and not tables:
            self.picture_block(s, box, imgs[0], n, sl)
            if ph is not None:
                remove_shape(ph)
            return
        if tables and not paras:
            self.table(s, box, tables[0], n, sl)
            if ph is not None:
                remove_shape(ph)
            return
        if ph is None:
            ph = s.shapes.add_textbox(*box)
            paras = plain(paras)
        fill_text(ph, paras, self.tpl, self.problems, self.where(n, sl))
        if imgs or tables:
            self.problems.append(("WARN", f"{self.where(n, sl)}: a column holds text and a picture/table; "
                                          f"only the text was placed (split with |||)"))

    def two_col(self, n, sl):
        blocks = sl["blocks"]
        if any(b["kind"] == "split" for b in blocks):
            i = next(k for k, b in enumerate(blocks) if b["kind"] == "split")
            left, right = blocks[:i], blocks[i + 1:]
        else:                                     # text + picture: text left, picture right …
            right = [b for b in blocks if b["kind"] in ("image", "table")][:1]
            left = [b for b in blocks if b not in right]
            eqs = [b for b in blocks if b["kind"] == "image" and "eq" in b.get("classes", [])]
            if eqs:                                          # equations go under the text, one per row
                return self.stacked(n, sl, [b for b in blocks if b not in eqs], eqs)
            if right and right[0]["kind"] == "image" and self.wide(right[0]):
                return self.stacked(n, sl, left, right)      # … unless the picture is wide: then below
            if self.variety and right and right[0]["kind"] == "image" and "two-col" not in sl["classes"]:
                self.flip = not self.flip                    # picture in its own half, alternating sides
                return self.split_image(n, sl, "right" if self.flip else "left", left, right[0])
        s = self.new("two")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        self.check_density(n, sl, self.paras_of(blocks))
        phs = sorted([ph for ph in s.placeholders if ph.placeholder_format.idx not in (0,)
                      and ph.placeholder_format.type is not None and int(ph.placeholder_format.type) in (2, 7)],
                     key=lambda p: p.left)
        if len(phs) >= 2:
            boxes = [(p.left, p.top, p.width, p.height) for p in phs[:2]]
        else:                  # no two-column layout: split the content area ourselves
            l, t, w, h = self.tpl.content_box
            g = int(w * 0.04)
            boxes = [(l, t, (w - g) // 2, h), (l + (w + g) // 2, t, (w - g) // 2, h)]
            if phs:            # the one body placeholder becomes the left column (keeps the theme's bullets)
                phs[0].left, phs[0].top, phs[0].width, phs[0].height = boxes[0]
                phs = [phs[0], None]
            else:
                phs = [None, None]
        for k, col in enumerate((left, right)):
            self.column(s, boxes[k], col, n, sl, phs[k] if k < len(phs) else None)
        self.finish(s, "two", sl["notes"])

    def wide(self, b):
        img = resolve_image(b["path"], self.sdir, self.cfg)
        if img is None:
            return False
        with Image.open(img) as im:
            return im.size[0] / im.size[1] > 2.0

    def stacked(self, n, sl, text_blocks, img_blocks):
        """Text on top, wide pictures/equations below at full width, one per row
        (a wide diagram in half a slide is unreadable)."""
        s = self.new("content")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        paras = self.paras_of(text_blocks)
        self.check_density(n, sl, paras)
        l, t, w, h = self.tpl.content_box
        body = next(ph for ph in s.placeholders if ph.placeholder_format.idx == 1)
        text_h = 0
        if paras:      # the content placeholder, cut to what the text needs: it keeps the theme's bullets
            need = text_height(paras, self.tpl, "+mn-lt", 1.0, w / EMU_PT - 14.4) / SAFETY + 7.2
            text_h = min(int(need * EMU_PT), int(h * 0.5))
            body.left, body.top, body.width, body.height = l, t, w, text_h
            fill_text(body, paras, self.tpl, self.problems, self.where(n, sl))
        gap = int(0.2 * 914400) if paras else 0
        # the pictures share what is left below the text, each centred in its row
        top, rest = t + text_h + gap, h - text_h - gap
        row = rest // len(img_blocks)
        for k, b in enumerate(img_blocks):
            self.picture_block(s, (l, top + k * row, w, row), b, n, sl)
        self.finish(s, "content", sl["notes"])

    def picture_block(self, s, box, b, n, sl):
        if b["path"].lower().endswith(".csv"):
            return self.chart(s, box, b, n, sl)
        img = resolve_image(b["path"], self.sdir, self.cfg)
        if img is None:
            self.problems.append(("FAIL", f"{self.where(n, sl)}: image {b['path']} not found "
                                          f"(looked in slides/, slides/figures/, report/, report/figures/)"))
            return
        place_picture(s, img, self.caption(s, box, b, n, sl), self.problems, self.where(n, sl),
                      natural_dpi=300 if "eq" in b.get("classes", []) else None)

    def caption(self, s, box, b, n, sl):
        l, t, w, h = box
        if not b["caption"]:
            return box
        cap_h = int(0.45 * 914400)
        tb = s.shapes.add_textbox(l, t + h - cap_h, w, cap_h)
        lv = {0: {"size": 16, "marL": 0, "spc": 0, "ln": 1.0}}
        fill_text(tb, [(0, runs(b["caption"], self.cites), False)], self.tpl, self.problems,
                  self.where(n, sl) + " caption", min_size=MIN_SMALL, levels=lv,
                  colour=self.ink()["muted"], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
        return (l, t, w, h - cap_h)

    def chart(self, s, box, b, n, sl):
        """CSV → native, editable PowerPoint chart. Error bars / log axis → matplotlib picture instead."""
        import deck_charts
        f = resolve_image(b["path"], self.sdir, self.cfg)
        if f is None:
            self.problems.append(("FAIL", f"{self.where(n, sl)}: chart data {b['path']} not found"))
            return
        box = self.caption(s, box, b, n, sl)
        data = deck_charts.read_csv(f)
        if deck_charts.needs_picture(data, b["attrs"]):
            P = self.tpl.palette
            png = deck_charts.matplotlib_png(data, b["attrs"], self.sdir / "_build", f"{f.stem}-{n}",
                                             family=self.tpl.fonts.minor, aspect=box[2] / box[3],
                                             colours=[P[f"accent{i}"] for i in range(1, 7)],
                                             ink=self.ink()["text"] if self.dark else None, bg=self.slide_bg)
            place_picture(s, png, box, self.problems, self.where(n, sl))
        else:
            deck_charts.native(s, box, data, b["attrs"], ink=self.ink()["text"] if self.dark else None)

    def image(self, n, sl):
        b = next(b for b in sl["blocks"] if b["kind"] == "image")
        eqs = [x for x in sl["blocks"] if x["kind"] == "image" and "eq" in x.get("classes", [])]
        if eqs and len(sl["blocks"]) > 1:                       # text first, then the equation(s)
            return self.stacked(n, sl, [x for x in sl["blocks"] if x not in eqs], eqs)
        s = self.new("title_only")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        l, t, w, h = self.tpl.content_box
        paras = [x for x in sl["blocks"] if x["kind"] == "para"]
        if paras:   # a one-line takeaway under the picture
            th = int(0.6 * 914400)
            tb = s.shapes.add_textbox(l, t + h - th, w, th)
            lv = {0: {"size": 20, "marL": 0, "spc": 0, "ln": 1.0}}
            fill_text(tb, plain(self.paras_of(paras)), self.tpl, self.problems, self.where(n, sl), min_size=16,
                      levels=lv, align=PP_ALIGN.CENTER)
            h -= th
        self.picture_block(s, (l, t, w, h), b, n, sl)
        self.finish(s, "title_only", sl["notes"])

    def table(self, s, box, b, n, sl):
        rows = b["rows"]
        if not rows:
            return
        ncol = max(len(r) for r in rows)
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        align = (b.get("align") or ["l"] * ncol) + ["l"] * ncol
        l, t, w, h = box
        cell_runs = [[runs(c, self.cites) for c in r] for r in rows]
        # column widths from content length, each at least 10% of the table
        lens = [max(max(sum(len(x) for x, _ in cr[c]) for cr in cell_runs), 3) for c in range(ncol)]
        tot = sum(lens)
        fr = [max(x / tot, 0.10) for x in lens]
        fr = [x / sum(fr) for x in fr]
        widths = [int(w * x) for x in fr]
        pad = 7.2 * 2                                            # cell margins left+right (pt)
        for size in range(20, MIN_SMALL - 1, -2):
            heights = []
            for r, cr in enumerate(cell_runs):
                lines = max(wrap_lines([(x, "b" if r == 0 else y) for x, y in c] or [("x", "")],
                                       self.tpl.fonts, "+mn-lt", size, widths[k] / EMU_PT - pad)
                            for k, c in enumerate(cr))
                heights.append(lines * size * 1.2 + 7.2)
            if sum(heights) <= h / EMU_PT * SAFETY:
                break
        else:
            self.problems.append(("FAIL", f"{self.where(n, sl)}: table too big for one slide even at "
                                          f"{MIN_SMALL} pt: drop rows/columns or split it"))
        total_h = int(sum(heights) * EMU_PT)
        if all(h1 <= size * 1.2 + 7.3 for h1 in heights):   # nothing wraps: shrink to the content, centred
            natural = [max(self.tpl.fonts.width("".join(x for x, _ in cr[c]), "+mn-lt", "b" if r == 0 else "",
                                                size) for r, cr in enumerate(cell_runs)) + pad + 24
                       for c in range(ncol)]
            tw = sum(natural) * EMU_PT
            if tw < w:
                widths = [int(x * EMU_PT * max(1.0, min(1.4, w * 0.75 / tw))) for x in natural]
                l += (w - sum(widths)) // 2
                w = sum(widths)
        t += max(0, (h - total_h) // 3)                       # upper third, not glued to the title
        shape = s.shapes.add_table(len(rows), ncol, l, t, w, total_h)
        tbl = shape.table
        for k, wd in enumerate(widths):
            tbl.columns[k].width = wd
        for r, hh in enumerate(heights):
            tbl.rows[r].height = int(hh * EMU_PT)
        al = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}
        for r, cr in enumerate(cell_runs):
            for c, rs in enumerate(cr):
                cell = tbl.cell(r, c)
                cell.margin_top = cell.margin_bottom = Pt(3.6)
                tf = cell.text_frame
                tf.word_wrap = True
                tf.clear()
                par = tf.paragraphs[0]
                par.alignment = al[align[c]]
                set_runs(par, [(x, ("b" + y) if r == 0 and "b" not in y else y) for x, y in rs], size)
        self.highlight_table(tbl, rows)
        return shape

    def highlight_table(self, tbl, rows):
        """The column that matters (Gap / Limitation) and our own row (ours / proposed) in the accent colour."""
        import deck_theme
        P = self.tpl.palette
        tint = deck_theme.mix(P["accent1"], "FFFFFF", 0.86)
        for c, head in enumerate(rows[0]):
            if re.search(r"\b(gap|gaps|limitations?|drawbacks?|research gap)\b", head, re.I):
                tbl.cell(0, c).fill.solid()
                tbl.cell(0, c).fill.fore_color.rgb = RGBColor.from_string(P["accent2"])
                for r in range(1, len(rows)):
                    tbl.cell(r, c).fill.solid()
                    tbl.cell(r, c).fill.fore_color.rgb = RGBColor.from_string(
                        deck_theme.mix(P["accent2"], "FFFFFF", 0.85))
        for r in range(1, len(rows)):
            if any(re.search(r"\b(ours|proposed|this work)\b", x, re.I) for x in rows[r]):
                for c in range(len(rows[r])):
                    tbl.cell(r, c).fill.solid()
                    tbl.cell(r, c).fill.fore_color.rgb = RGBColor.from_string(tint)
                    for p in tbl.cell(r, c).text_frame.paragraphs:
                        for run in p.runs:
                            run.font.bold = True

    def table_slide(self, n, sl):
        s = self.new("title_only")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        b = next(b for b in sl["blocks"] if b["kind"] == "table")
        l, t, w, h = self.tpl.content_box
        paras = [x for x in sl["blocks"] if x["kind"] in ("para", "bullets")]
        if paras:
            th = int(1.0 * 914400)
            tb = s.shapes.add_textbox(l, t + h - th, w, th)
            fill_text(tb, plain(self.paras_of(paras)), self.tpl, self.problems, self.where(n, sl), min_size=16,
                      anchor=MSO_ANCHOR.BOTTOM)
            h -= th
        self.table(s, (l, t, w, h), b, n, sl)
        self.finish(s, "title_only", sl["notes"])

    def results(self, n, sl):
        """Big numbers: '- 0.92 :: FedProx accuracy on non-IID clients'."""
        if any(b["kind"] == "image" for b in sl["blocks"]):          # numbers + the chart they come from
            import deck_patterns
            return deck_patterns.numbers_chart(self, n, sl)
        s = self.new("title_only")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        cards, rest = [], []
        for b in sl["blocks"]:
            if b["kind"] == "bullets":
                for lvl, t in b["items"]:
                    if "::" in t:
                        v, lab = t.split("::", 1)
                        cards.append((v.strip(), lab.strip()))
                    else:
                        rest.append({"kind": "bullets", "items": [(lvl, t)]})
            else:
                rest.append(b)
        if not cards:
            self.problems.append(("FAIL", f"{self.where(n, sl)}: a .results slide needs lines like "
                                          f"“- 0.92 :: FedProx accuracy”"))
        l, t, w, h = self.tpl.content_box
        if rest:
            th = int(1.3 * 914400)
            tb = s.shapes.add_textbox(l, t + h - th, w, th)
            fill_text(tb, plain(self.paras_of(rest)), self.tpl, self.problems, self.where(n, sl), min_size=16,
                      anchor=MSO_ANCHOR.BOTTOM)
            h -= th
        per_row = min(len(cards), 4) or 1
        nrows = (len(cards) + per_row - 1) // per_row
        gap = int(0.3 * 914400)
        cw = (w - gap * (per_row - 1)) // per_row
        ch = min(int(2.3 * 914400), (h - gap * (nrows - 1)) // max(nrows, 1))
        top0 = t + (h - (ch * nrows + gap * (nrows - 1))) // 2
        ink = self.ink()
        for k, (v, lab) in enumerate(cards):
            r, c = divmod(k, per_row)
            x, y = l + c * (cw + gap), top0 + r * (ch + gap)
            card = self.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, (x, y, cw, ch), ink["card"])
            card.adjustments[0] = 0.08
            vb = s.shapes.add_textbox(x, y + int(ch * 0.10), cw, int(ch * 0.50))
            lv = {0: {"size": 48, "marL": 0, "spc": 0, "ln": 1.0}}
            fill_text(vb, [(0, [(v, "b")], False)], self.tpl, self.problems, self.where(n, sl) + " number",
                      family="+mj-lt", min_size=24, levels=lv, colour=ink["number"], align=PP_ALIGN.CENTER,
                      anchor=MSO_ANCHOR.BOTTOM)
            lb = s.shapes.add_textbox(x, y + int(ch * 0.60), cw, int(ch * 0.36))
            lv = {0: {"size": 18, "marL": 0, "spc": 0, "ln": 1.0}}
            fill_text(lb, [(0, runs(lab, self.cites), False)], self.tpl, self.problems,
                      self.where(n, sl) + " label", min_size=MIN_SMALL, levels=lv, colour=ink["text"],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
        self.finish(s, "title_only", sl["notes"])

    def references(self, n, sl, bib):
        s = self.new("content")
        title_text(s.shapes.title, sl["title"] or "References", self.tpl, self.cites, self.problems,
                   self.where(n, sl))
        body = next(ph for ph in s.placeholders if ph.placeholder_format.idx == 1)
        paras = []
        for k, key in enumerate(self.cites.order, 1):
            paras.append((0, [(f"[{k}]  ", "b")] + ref_runs(bib[key]), False))
        if not paras:
            paras = [(0, [("No citations in the slides yet.", "i")], False)]
        lv = {0: {"size": 18, "marL": 0, "spc": 8, "ln": 1.0}}
        fill_text(body, paras, self.tpl, self.problems, self.where(n, sl), min_size=MIN_SMALL, levels=lv)
        self.finish(s, "content", sl["notes"])

    def qa(self, n, sl):
        """Thank you / Questions, in the cover's design (the talk ends where it began)."""
        W, H = self.tpl.W, self.tpl.H
        text = sl["title"] if sl else "Thank you"
        where = "Q&A slide"
        if str(self.cfg.get("cover") or "theme") == "college":
            s = self.cover_base()
            ink, P = self.ink(), self.tpl.palette
            if self.logo:
                pic = place_picture(s, self.logo, (int(W / 2 - 0.9 * IN), int(0.3 * IN), int(1.8 * IN),
                                                   int(1.25 * IN)), self.problems, where, valign="top")
                self.logo_plate_behind(s, pic)
            t = s.shapes.title
            t.left, t.top, t.width, t.height = int(W * 0.1), int(H * 0.34), int(W * 0.8), int(1.4 * IN)
            title_text(t, text, self.tpl, self.cites, self.problems, where, size=48, min_size=28,
                       align=PP_ALIGN.CENTER, colour=ink["accent"])
            t.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
            self.shape(s, MSO_SHAPE.RECTANGLE, (W * 0.42, H * 0.34 + 1.55 * IN, W * 0.16, 0.035 * IN), ink["accent"])
            self.text(s, (int(W * 0.1), int(H * 0.34 + 1.75 * IN), int(W * 0.8), int(0.6 * IN)),
                      [(0, [("Questions?", "")])], 24, 0, {"title": where, "line": 0}, colour=ink["muted"],
                      align=PP_ALIGN.CENTER, min_size=16)
            self.finish(s, "title_only", sl["notes"] if sl else [], footer="number")
            return
        s = self.new("title")
        title_text(s.shapes.title, text, self.tpl, self.cites, self.problems, where, min_size=28)
        sub = next((ph for ph in s.placeholders if ph.placeholder_format.idx == 1), None)
        if sub is not None:
            size = min(font_size_of(sub, self.tpl, 20), 24)
            fill_text(sub, [(0, [("Questions?", "")], False)], self.tpl, self.problems, where, min_size=14,
                      levels={0: {"size": size, "marL": 0, "spc": 0, "ln": 1.0}})
        if self.logo and self.cfg.get("logo", True) and not self.tpl.has_logo:
            box = (int(W - W * 0.05 - 1.4 * IN), int(H * 0.06), int(1.4 * IN), int(1.4 * IN))
            place_picture(s, self.logo, box, self.problems, where, valign="top")
        self.finish(s, "title", sl["notes"] if sl else [], footer=False)

    # ── livelier slide types: cards, steps, timeline, statement, quote, picture beside text ──
    def items_of(self, sl):
        """Top-level list items as [text, [sub-item texts]], and the blocks that aren't lists."""
        items, rest = [], []
        for b in sl["blocks"]:
            if b["kind"] == "bullets":
                for lvl, t in b["items"]:
                    if lvl == 0 or not items:
                        items.append([t, []])
                    else:
                        items[-1][1].append(t)
            else:
                rest.append(b)
        return items, rest

    def takeaway(self, s, rest, box, n, sl):
        """A line of text under a diagram-like slide; → height left above it."""
        l, t, w, h = box
        paras = plain(self.paras_of(rest))
        if not paras:
            return h
        th = int(1.0 * IN)
        tb = s.shapes.add_textbox(l, t + h - th, w, th)
        fill_text(tb, paras, self.tpl, self.problems, self.where(n, sl), min_size=16, anchor=MSO_ANCHOR.BOTTOM,
                  colour=self.ink()["text"])
        return h - th - int(0.15 * IN)

    def uniform(self, filled):
        """Boxes of one kind (card bodies, step texts) get one size: the smallest any of them needed."""
        smallest = min((got for *_, got in filled), default=None)
        for shape, args, kw, got in filled:
            if got != smallest:
                fill_text(shape, *args, **{**kw, "max_size": smallest})

    def fallback(self, n, sl, why):
        self.problems.append(("WARN", f"{self.where(n, sl)}: {why}; drawn as bullets"))
        return self.bullets(n, sl)

    def cards(self, n, sl):
        """Each top-level item in its own card: '- **Privacy**: data stays on the device'."""
        items, rest = self.items_of(sl)
        if not 2 <= len(items) <= 6:
            return self.fallback(n, sl, f".cards takes 2–6 items, not {len(items)}")
        s = self.new("title_only")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        l, t, w, h = self.tpl.content_box
        h = self.takeaway(s, rest, (l, t, w, h), n, sl)
        ink, k = self.ink(), len(items)
        long = sum(len(x[0].split()) for x in items) / k > 9
        cols = {2: 2, 3: 3, 4: 2 if long else 4, 5: 3, 6: 3}[k]
        rows = (k + cols - 1) // cols
        gap = int(0.3 * IN)
        cw = (w - gap * (cols - 1)) // cols
        pad, badge = int(0.25 * IN), int(0.6 * IN)
        import deck_patterns
        parts, icons = [], []
        for icon, head, body, subs in deck_patterns.items_with_icons(self, items):
            body_paras = ([(0, runs(body, self.cites))] if body else []) + [(1, runs(x, self.cites)) for x in subs]
            parts.append((head, body_paras))
            icons.append(icon)
        # card height: what the longest card needs at 20 pt, within the space there is
        lv = {0: {"size": 20, "marL": 0, "spc": 0, "ln": 1.0}, 1: {"size": 18, "marL": 0, "spc": 5, "ln": 1.0}}
        need = 0
        for head, bp in parts:
            bw = (cw - 2 * pad) if head else (cw - 2 * pad - badge - int(0.15 * IN))
            hh = text_height([(lvl, rs, False) for lvl, rs in bp], self.tpl, "+mn-lt", 1.0, bw / EMU_PT - 14.4, lv)
            need = max(need, (hh / SAFETY + 7.2) * EMU_PT + (badge + pad * 2 if head else pad * 2))
        ch = int(min((h - gap * (rows - 1)) // rows, max(need + 0.3 * IN, (2.6 if rows == 1 else 2.0) * IN)))
        top0 = t + (h - (ch * rows + gap * (rows - 1))) // 3
        bodies = []
        for i, (head, bp) in enumerate(parts):
            r, c = divmod(i, cols)
            in_row = min(cols, k - r * cols)
            x0 = l + (w - (in_row * cw + (in_row - 1) * gap)) // 2       # a shorter last row is centred
            x, y = x0 + c * (cw + gap), top0 + r * (ch + gap)
            card = self.shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, (x, y, cw, ch), ink["card"])
            card.adjustments[0] = 0.07
            deck_patterns.marker(self, s, "icon" if icons[i] else "number", (x + pad, y + pad, badge, badge), i,
                                 icons[i])
            if head:
                self.text(s, (x + pad + badge + int(0.15 * IN), y + pad - int(0.05 * IN),
                              cw - 2 * pad - badge - int(0.15 * IN), badge + int(0.1 * IN)),
                          [(0, [(x, st + "b" if "b" not in st else st) for x, st in runs(head, self.cites)])],
                          22, n, sl, colour=ink["title"], family="+mj-lt",
                          anchor=MSO_ANCHOR.MIDDLE, min_size=16, what=" card heading")
                box = (x + pad, y + pad + badge + int(0.12 * IN), cw - 2 * pad, ch - 2 * pad - badge - int(0.12 * IN))
            else:
                box = (x + pad + badge + int(0.15 * IN), y + pad, cw - 2 * pad - badge - int(0.15 * IN), ch - 2 * pad)
            if bp:
                tb = s.shapes.add_textbox(*(int(v) for v in box))
                args = ([(lvl, rs, False) for lvl, rs in bp], self.tpl, self.problems, self.where(n, sl) + " card")
                kw = dict(min_size=14, levels=lv, colour=ink["text"], anchor=MSO_ANCHOR.TOP)
                bodies.append((tb, args, kw, fill_text(tb, *args, **kw)))
        self.uniform(bodies)
        self.finish(s, "title_only", sl["notes"])

    def steps(self, n, sl):
        """A process, left to right: numbered arrows with the text under each."""
        import deck_theme
        items, rest = self.items_of(sl)
        if not 2 <= len(items) <= 6:
            return self.fallback(n, sl, f".steps takes 2–6 items, not {len(items)}")
        s = self.new("title_only")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        l, t, w, h = self.tpl.content_box
        h = self.takeaway(s, rest, (l, t, w, h), n, sl)
        ink, P, k = self.ink(), self.tpl.palette, len(items)
        gap, ah = int(0.06 * IN), int(1.0 * IN)
        aw = (w - gap * (k - 1)) // k
        parts = [split_head(text) + (subs,) for text, subs in items]
        lv = {0: {"size": 20, "marL": 0, "spc": 0, "ln": 1.0}, 1: {"size": 18, "marL": 0, "spc": 5, "ln": 1.0}}
        need = max(text_height([(0, runs(body, self.cites), False)] + [(1, runs(x, self.cites), False) for x in subs],
                               self.tpl, "+mn-lt", 1.0, aw / EMU_PT - 30, lv) / SAFETY + 7.2
                   for head, body, subs in parts) * EMU_PT
        group = ah + int(0.25 * IN) + min(need, h - ah - int(0.25 * IN))
        y = t + max(0, (h - group) // 3)
        start = P["accent1"] if not self.dark else deck_theme.mix(P["accent1"], "FFFFFF", 0.25)
        end = deck_theme.mix(start, P["dk2"] if not self.dark else self.slide_bg, 0.55 if not self.dark else 0.35)
        bodies = []
        for i, (head, body, subs) in enumerate(parts):
            x = l + i * (aw + gap)
            colour = deck_theme.mix(start, end, i / max(k - 1, 1))
            kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
            a = self.shape(s, kind, (x, y, aw + (int(0.25 * IN) if i < k - 1 else 0), ah), colour)
            a.adjustments[0] = 0.3
            tf = a.text_frame
            tf.margin_left, tf.margin_right = int(0.32 * IN) if i else int(0.12 * IN), int(0.3 * IN)
            label = head or f"{i + 1:02d}"
            self.text(s, None, [(0, [(label, "b")])], 20 if head else 26, n, sl, colour="FFFFFF", family="+mj-lt",
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, shape=a, min_size=12, what=" step")
            paras = ([(0, runs(body, self.cites), False)] if body else []) + \
                [(1, runs(x2, self.cites), False) for x2 in subs]
            if paras:
                tb = s.shapes.add_textbox(x + int(0.05 * IN), y + ah + int(0.2 * IN), aw - int(0.1 * IN),
                                          t + h - (y + ah + int(0.2 * IN)))
                args = (paras, self.tpl, self.problems, self.where(n, sl) + " step text")
                kw = dict(min_size=14, levels=lv, colour=ink["text"], anchor=MSO_ANCHOR.TOP)
                bodies.append((tb, args, kw, fill_text(tb, *args, **kw)))
        self.uniform(bodies)
        self.finish(s, "title_only", sl["notes"])

    def timeline(self, n, sl, arrow=False):
        """Dates or phases on a line: '- 2024 :: Data collected' (or '- **Phase 1**: …')."""
        items, rest = self.items_of(sl)
        if not 2 <= len(items) <= 7:
            return self.fallback(n, sl, f".timeline takes 2–7 items, not {len(items)}")
        s = self.new("title_only")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        l, t, w, h = self.tpl.content_box
        h = self.takeaway(s, rest, (l, t, w, h), n, sl)
        ink, k = self.ink(), len(items)
        mid = t + int(h * 0.36)
        self.shape(s, MSO_SHAPE.RECTANGLE, (l, mid - int(0.025 * IN), w - (int(0.2 * IN) if arrow else 0),
                                             int(0.05 * IN)), ink["line"])
        if arrow:                                   # a roadmap points forward
            a = self.shape(s, MSO_SHAPE.ISOSCELES_TRIANGLE, (l + w - int(0.3 * IN), mid - int(0.14 * IN),
                                                            int(0.28 * IN), int(0.3 * IN)), ink["line"])
            a.rotation = 90
        iw = w // k - int(0.12 * IN)
        dot = int(0.34 * IN)
        bodies = []
        for i, (text, subs) in enumerate(items):
            if "::" in text:
                head, body = (x.strip() for x in text.split("::", 1))
            else:
                head, body = split_head(text)
                if arrow and not head:
                    head = f"{i + 1}"
            cx = l + int(w * (i + 0.5) / k)
            self.shape(s, MSO_SHAPE.OVAL, (cx - dot // 2, mid - dot // 2, dot, dot), ink["accent"])
            if head:
                self.text(s, (cx - iw // 2, mid - int(1.05 * IN), iw, int(0.8 * IN)), [(0, runs(head, self.cites))],
                          22, n, sl, colour=ink["number"], family="+mj-lt", align=PP_ALIGN.CENTER,
                          anchor=MSO_ANCHOR.BOTTOM, min_size=14, what=" timeline label")
            paras = ([(0, runs(body, self.cites), False)] if body else []) + \
                [(0, runs(x2, self.cites), False) for x2 in subs]
            if paras:
                tb = s.shapes.add_textbox(cx - iw // 2, mid + int(0.35 * IN), iw, t + h - mid - int(0.35 * IN))
                args = (paras, self.tpl, self.problems, self.where(n, sl) + " timeline text")
                kw = dict(min_size=14, levels={0: {"size": 18, "marL": 0, "spc": 4, "ln": 1.0}},
                          colour=ink["text"], anchor=MSO_ANCHOR.TOP, align=PP_ALIGN.CENTER)
                bodies.append((tb, args, kw, fill_text(tb, *args, **kw)))
        self.uniform(bodies)
        self.finish(s, "title_only", sl["notes"])

    def statement(self, n, sl):
        """One message, large, on a coloured slide (the title is the message; a line of text may follow)."""
        s = self.new("title_only", look="accent")
        W, H = self.tpl.W, self.tpl.H
        ink = self.ink()
        paras = plain(self.paras_of(sl["blocks"]))
        t = s.shapes.title
        t.left, t.top, t.width, t.height = int(W * 0.1), int(H * 0.16), int(W * 0.8), int(H * 0.44)
        title_text(t, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl), size=40, min_size=28,
                   align=PP_ALIGN.CENTER, colour=ink["title"])
        t.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM if paras else MSO_ANCHOR.MIDDLE
        if paras:
            self.shape(s, MSO_SHAPE.RECTANGLE, (W / 2 - 0.6 * IN, H * 0.63, 1.2 * IN, 0.06 * IN),
                       self.tpl.palette["accent2"])
            tb = s.shapes.add_textbox(int(W * 0.15), int(H * 0.68), int(W * 0.7), int(H * 0.2))
            fill_text(tb, paras, self.tpl, self.problems, self.where(n, sl), min_size=16,
                      levels={0: {"size": 22, "marL": 0, "spc": 6, "ln": 1.0}}, colour=ink["text"],
                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.TOP)
        self.finish(s, "title_only", sl["notes"], footer="number")

    def quote(self, n, sl):
        """A quotation: the paragraph is the quote, a line starting with — is who said it."""
        s = self.new("title_only")
        title_text(s.shapes.title, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl))
        l, t, w, h = self.tpl.content_box
        ink = self.ink()
        quote, who = [], []
        for b in sl["blocks"]:
            if b["kind"] == "para":
                txt = b["text"]
                if re.match(r"^(—|--|–|~)\s*", txt):
                    who.append(re.sub(r"^(—|--|–|~)\s*", "— ", txt))
                else:
                    quote.append(txt)
        self.text(s, (l - int(0.1 * IN), t - int(0.3 * IN), int(1.4 * IN), int(2.2 * IN)), [(0, [("“", "b")])], 110, n,
                  sl, colour=ink["accents"][1], family="+mj-lt", min_size=60, what=" quote mark")
        qx = l + int(1.3 * IN)
        qh = int(h * 0.7)
        if quote:
            rs = [[(x, (st + "i") if st not in ("code", "todo") and "i" not in st else st)
                   for x, st in runs(q, self.cites)] for q in quote]
            self.text(s, (qx, t + int(0.3 * IN), w - int(1.6 * IN), qh), [(0, r) for r in rs], 30, n, sl,
                      colour=ink["text"], family="+mj-lt", min_size=20, anchor=MSO_ANCHOR.MIDDLE, what=" quote")
        if who:
            self.text(s, (qx, t + int(0.3 * IN) + qh, w - int(1.6 * IN), int(0.6 * IN)),
                      [(0, runs(x, self.cites)) for x in who], 20, n, sl, colour=ink["muted"],
                      align=PP_ALIGN.RIGHT, min_size=14, what=" attribution")
        self.finish(s, "title_only", sl["notes"])

    def split_image(self, n, sl, side, text_blocks, pic):
        """Picture in one half of the slide (on a tinted panel, fitted, never cropped), title and text in the other."""
        s = self.new("content")
        s._element.set("showMasterSp", "0")              # the theme's bands would run across the picture
        W, H = self.tpl.W, self.tpl.H
        ink = self.ink()
        half = W // 2
        px = 0 if side == "left" else W - half
        panel = self.shape(s, MSO_SHAPE.RECTANGLE, (px, 0, half, H), ink["card"])
        to_back(panel)
        self.picture_block(s, (px + int(0.45 * IN), int(0.45 * IN), half - int(0.9 * IN), H - int(0.9 * IN)), pic, n, sl)
        m = int(0.6 * IN)
        tx = half + m if side == "left" else m
        tw = W - half - 2 * m
        t = s.shapes.title
        t.left, t.top, t.width, t.height = tx, int(0.45 * IN), tw, int(1.5 * IN)
        title_text(t, sl["title"], self.tpl, self.cites, self.problems, self.where(n, sl), colour=ink["title"],
                   size=min(font_size_of(t, self.tpl, 32), 30))
        t.text_frame.vertical_anchor = MSO_ANCHOR.BOTTOM
        self.shape(s, MSO_SHAPE.RECTANGLE, (tx, int(2.05 * IN), int(1.0 * IN), int(0.06 * IN)),
                   self.tpl.palette["accent2"])
        body = next(ph for ph in s.placeholders if ph.placeholder_format.idx == 1)
        paras = self.paras_of(text_blocks)
        self.check_density(n, sl, paras)
        body.left, body.top, body.width, body.height = tx, int(2.3 * IN), tw, H - int(2.3 * IN) - int(0.75 * IN)
        if paras:
            fill_text(body, paras, self.tpl, self.problems, self.where(n, sl))
            if self.tpl.is_dark(self.tpl.layout["content"]) != self.dark:
                for p in body.text_frame.paragraphs:
                    for r in p.runs:
                        r.font.color.rgb = RGBColor.from_string(ink["text"])
        self.finish(s, "content", sl["notes"], footer="number")

    def picture_side(self, n, sl, side):
        pics = [b for b in sl["blocks"] if b["kind"] == "image"]
        if not pics:
            return self.fallback(n, sl, f".image-{side} needs a picture")
        if self.wide(pics[0]):
            self.problems.append(("WARN", f"{self.where(n, sl)}: {pics[0]['path']} is wide: in half a slide it "
                                          f"is hard to read (without .image-{side} it goes full width)"))
        return self.split_image(n, sl, side, [b for b in sl["blocks"] if b is not pics[0]], pics[0])


def plain(paras):
    """Text boxes (unlike placeholders) don't inherit the theme's bullets: write them as plain lines."""
    return [(lvl, rs, False) for lvl, rs, _ in paras]


def clean_bib(s):
    s = re.sub(r"\\[\"'`^~=.]\{?(\w)\}?", r"\1", str(s))
    return re.sub(r"[{}]", "", s).strip()


def ref_runs(e):
    """Short IEEE-like reference: A. Author et al., “Title,” Venue, Year."""
    authors = [a.strip() for a in clean_bib(e.get("author", "")).split(" and ") if a.strip()]

    def short(a):
        if "," in a:
            fam, giv = [x.strip() for x in a.split(",", 1)]
        else:
            parts = a.split()
            fam, giv = parts[-1], " ".join(parts[:-1])
        ini = " ".join(g[0] + "." for g in giv.split() if g)
        return f"{ini} {fam}".strip()

    who = short(authors[0]) + (" et al." if len(authors) > 2 else
                               f" and {short(authors[1])}" if len(authors) == 2 else "") if authors else ""
    venue = clean_bib(e.get("booktitle") or e.get("journal") or e.get("publisher") or "")
    venue = re.sub(r"^Proceedings of (the )?", "Proc. ", venue)
    out = [(f"{who}, " if who else "", ""), (f"“{clean_bib(e.get('title', ''))},” ", "")]
    if venue:
        out.append((venue, "i"))
    out.append((f", {e.get('year', '')}." if e.get("year") else ".", ""))
    return [(t, s) for t, s in out if t]


# ── render: PDF + contact sheet + overflow check on the real rendering ──
def render_pdf(pptx: Path, out_dir: Path):
    if not shutil.which("soffice"):
        return None, "LibreOffice not found: no PDF/preview (the .pptx is fine)"
    with tempfile.TemporaryDirectory() as prof:
        r = subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", "pdf",
                            "--outdir", str(out_dir), str(pptx)], capture_output=True, text=True, timeout=240)
    pdf = out_dir / (pptx.stem + ".pdf")
    if r.returncode != 0 or not pdf.exists():
        return None, f"LibreOffice could not convert the deck: {r.stderr.strip()[:200]}"
    return pdf, None


def rendered_overflow(pdf: Path, footer_top_pt: float, footer_text=""):
    """Text drawn below the footer line or off the slide in the real rendering."""
    import pymupdf
    out = []
    doc = pymupdf.open(pdf)
    for i, page in enumerate(doc, 1):
        W, H = page.rect.width, page.rect.height
        for b in page.get_text("dict")["blocks"]:
            for line in b.get("lines", []):
                for sp in line["spans"]:
                    txt = sp["text"].strip()
                    if not txt or txt.isdigit() or (footer_text and txt in footer_text):
                        continue                          # footer and slide number live down there
                    x0, y0, x1, y1 = sp["bbox"]
                    off = x1 > W + 1 or y1 > H + 1 or x0 < -1
                    low = sp["size"] > 12.5 and y1 > footer_top_pt + 2
                    if off or low:
                        out.append(("FAIL", f"slide {i}: “{sp['text'].strip()[:40]}” is drawn "
                                            f"{'off the slide' if off else 'over the footer'} in the rendering"))
                        break
                else:
                    continue
                break
    return out


def contact_sheet(pdf: Path, dest: Path, cols=4, dpi=40):
    import pymupdf
    doc = pymupdf.open(pdf)
    ims = []
    for p in doc:
        pix = p.get_pixmap(dpi=dpi)
        ims.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
    w, h = ims[0].size
    pad = 6
    rows = (len(ims) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (w + pad) + pad, rows * (h + pad) + pad), "#777")
    for k, im in enumerate(ims):
        sheet.paste(im, (pad + (k % cols) * (w + pad), pad + (k // cols) * (h + pad)))
    sheet.save(dest)


# ── build ────────────────────────────────────────────────────
def load_bib(sdir, cfg):
    import cite
    for d in (sdir, cfg["_report"]):
        if d and (d / "refs.bib").exists():
            return {e["ID"]: e for e in cite.load_bib(d).entries}
    return {}


def auto_kind(sl, kind, recent):
    """variety: auto. Plain bullet slides that read better another way become cards, steps or a statement,
    so a deck isn't fifteen identical bullet slides. Forcing a type ({.bullets}) turns this off per slide."""
    if kind != "bullets" or set(sl["classes"]) & set(TYPES):
        return kind
    lists = [b for b in sl["blocks"] if b["kind"] == "bullets"]
    paras = [b for b in sl["blocks"] if b["kind"] == "para"]
    others = [b for b in sl["blocks"] if b["kind"] not in ("bullets", "para")]
    if others:
        return kind
    if not lists and len(paras) == 1 and len(paras[0]["text"].split()) <= 25 and len(sl["title"].split()) <= 12:
        return "statement"
    if len(lists) != 1 or len(paras) > 1 or any(lvl for lvl, _ in lists[0]["items"]):
        return kind
    top = [t for _, t in lists[0]["items"]]
    words = [len(t.split()) for t in top]
    if lists[0]["ordered"] and 3 <= len(top) <= 5 and max(words) <= 16:
        return "steps"
    if 2 <= len(top) <= 6 and all(split_head(t)[0] for t in top) and max(words) <= 30:
        return "cards"
    return kind


def assemble(sdir, cfg, slides, problems):
    """slides (parsed slides.md) → Builder with the finished presentation in b.prs."""
    tpath = template_path(sdir, cfg)
    tpl = Template(tpath, problems)
    bib = load_bib(sdir, cfg)
    cites = Cites(bib)
    for sl in slides:                       # $$ equations $$ → pictures, placed like any image
        for blk in sl["blocks"]:
            if blk["kind"] == "equation":
                png = render_equation(blk["tex"], sdir / "_build")
                blk.update(kind="image", path=str(png), caption="", classes=["eq"], attrs={})
    say(f"building {len(slides)} slide(s) from slides.md with {tpath.name}")
    import deck_diagrams
    if cfg["_report"]:                                  # the report's figures are reused on slides
        import report
        report.render_diagrams(cfg["_report"])
    deck_diagrams.render(sdir, cfg, family=tpl.fonts.minor)
    b = Builder(sdir, cfg, tpl, cites, problems)

    main = [s for s in slides if "backup" not in s["classes"]]
    backup = [s for s in slides if "backup" in s["classes"]]
    has_qa = any(slide_kind(s) == "qa" for s in main)

    import deck_patterns
    b.title_slide()
    n, sec = 1, 0
    b.sections = [s["title"] for s in main if s["level"] == 1]
    want = str(cfg.get("agenda", "auto")).lower()
    if want in ("true", "yes", "on") or (want == "auto" and len(b.sections) >= 3):
        n += 1
        deck_patterns.agenda_slide(b, n, b.sections)
        b.progress = True
    deferred_refs = None
    recent, auto = [], []

    def one(sl, n):
        nonlocal sec, deferred_refs
        b.cur = sl
        k = slide_kind(sl)
        if b.variety and not set(sl["classes"]) & set(TYPES):
            d, why = deck_patterns.choose(sl, k)
            if not d:
                d = auto_kind(sl, k, recent)
                why = {"cards": "labelled items", "steps": "a short numbered list", "statement": "one sentence"}.get(d)
            if d and d != k:
                auto.append(f"{n} {d.replace('_', '-')}: {why}")
                k = d
        recent.append(k)
        if k == "section":
            sec += 1
            b.section(n, sl, label=f"Part {sec}", current=sec - 1 if b.progress else None)
        elif k == "references":
            deferred_refs = (n, sl, len(b.prs.slides._sldIdLst))   # filled after all citations are numbered
            b.new("content")
        elif k in ("image-left", "image-right"):
            b.picture_side(n, sl, k.split("-")[1])
        else:
            P = deck_patterns
            draw = {"qa": b.qa, "two-col": b.two_col, "image": b.image, "table": b.table_slide,
                    "results": b.results, "cards": b.cards, "steps": b.steps, "timeline": b.timeline,
                    "statement": b.statement, "quote": b.quote,
                    "lead": lambda n, sl: P.lead(b, n, sl), "lead-cards": lambda n, sl: P.lead(b, n, sl, cards=True),
                    "numbered": lambda n, sl: P.numbered(b, n, sl),
                    "checklist": lambda n, sl: P.numbered(b, n, sl, check=True),
                    "compare": lambda n, sl: P.compare(b, n, sl), "proscons": lambda n, sl: P.proscons(b, n, sl),
                    "columns": lambda n, sl: P.columns(b, n, sl), "callouts": lambda n, sl: P.callouts(b, n, sl),
                    "insight": lambda n, sl: P.insight(b, n, sl),
                    "numbers-chart": lambda n, sl: P.numbers_chart(b, n, sl),
                    "definition": lambda n, sl: P.definition(b, n, sl),
                    "takeaways": lambda n, sl: P.takeaways(b, n, sl),
                    "roadmap": lambda n, sl: b.timeline(n, sl, arrow=True)}
            draw.get(k.replace("_", "-"), b.bullets)(n, sl)
        b.cur = None

    for sl in main:
        n += 1
        one(sl, n)
    if not has_qa and cfg.get("qa", True):
        n += 1
        b.qa(n, None)
    if backup:
        n += 1
        b.section(n, {"title": "Backup slides", "notes": [], "line": 0}, label="For questions")
        for sl in backup:
            n += 1
            one(sl, n)
    if deferred_refs:
        rn, rsl, idx = deferred_refs
        b.cur = rsl
        b.references(rn, rsl, bib)                     # appended at the end …
        b.cur = None
        lst = b.prs.slides._sldIdLst                   # … then moved into the placeholder's position
        new = lst[-1]
        lst.remove(new)
        old = lst[idx]
        lst.insert(idx, new)
        b.prs.part.drop_rel(old.rId)
        lst.remove(old)
    elif cites.order:
        problems.append(("WARN", "slides cite papers but have no {.references} slide"))
    for k in sorted(cites.missing):
        problems.append(("FAIL", f"[@{k}] is cited in slides.md but not in refs.bib"))
    if auto:
        problems.append(("INFO", "designs chosen from the content (force one with {.bullets}, {.compare} …; "
                                 "variety: off keeps plain slides):\n      " + "\n      ".join(auto)))
    return b


def build(sdir: Path, final=False):
    cfg = load_config(sdir)
    problems = []
    slides = parse((sdir / "slides.md").read_text())
    b = assemble(sdir, cfg, slides, problems)
    tpl = b.tpl
    out = sdir / "out"
    build_dir = sdir / "_build"
    out.mkdir(exist_ok=True)
    build_dir.mkdir(exist_ok=True)
    pptx = out / "slides.pptx"
    b.prs.save(pptx)
    say(f"wrote {pptx.relative_to(sdir.parent)} ({len(b.prs.slides._sldIdLst)} slides)")
    pdf, err = render_pdf(pptx, out)
    if err:
        problems.append(("WARN", err))
    else:
        footer_top = min((ph.top for ph in tpl.layout["content"].placeholders
                          if ph.placeholder_format.type is not None and int(ph.placeholder_format.type) in (13, 15, 16)),
                         default=tpl.H) / EMU_PT
        problems += rendered_overflow(pdf, footer_top, b.footer)
        contact_sheet(pdf, build_dir / "slides.png")
        problems.append(("INFO", f"slide overview: {(build_dir / 'slides.png').relative_to(sdir.parent)}"))
    import deck_check
    problems += deck_check.run(sdir, cfg, slides, final)
    return print_problems(problems, final)


def print_problems(problems, final):
    print()
    seen = set()
    for level, msg in problems:
        if (level, msg) in seen:
            continue
        seen.add((level, msg))
        mark = {"FAIL": "\033[31m✗\033[0m", "WARN": "\033[33m!\033[0m", "INFO": "·"}[level]
        print(f"  {mark} {msg}")
    fails = {p for p in problems if p[0] == "FAIL"}
    warns = {p for p in problems if p[0] == "WARN"}
    print(f"\n  {len(fails)} failure(s), {len(warns)} warning(s)"
          + ("" if final else "  [draft build: use --final before presenting]"))
    return not fails


# ── init ─────────────────────────────────────────────────────
STARTER_YML = """\
# Slide deck settings. `report slides themes` shows every style and palette.
title: "{title}"
# subtitle: "Final review"
theme:
  style: clean            # clean | band | sidebar | minimal | bold | soft | geometric
  palette: ocean          # ocean royal teal forest mint sunset crimson berry plum slate earth gold mono
  # primary: "#7B2CBF"    # or your own colour (the rest is derived from it)
  fonts: classic          # classic | serif-titles | formal | plain | times
cover: college            # first slide: logo + college, title, presented by / guide (me.yml)
# cover: first-page.pptx  # or your college's first slide with {{{{title}}}} {{{{name}}}} {{{{guide}}}} {{{{date}}}}
# logo_on_slides: true    # small logo on every slide
# team:                   # more than one presenter
#   - {{name: "A Student", roll: "101"}}
# variety: off            # keep plain bullet slides as bullets
# template: college.pptx  # a full college template instead of a theme (see template-extract)
talk_minutes: 15          # used to warn about too many / too few slides
# footer: "Short title"   # default: the title
# qa: false               # no automatic "Thank you / Questions?" slide
"""

STARTER_MD = """\
<!--
One "## Title" per slide; "# Title" makes a section divider.
Write the title as the slide's claim ("FedProx beats FedAvg on non-IID data"), not a label ("Results").
The type is picked from the content, or forced with {.results} {.two-col} {.image} {.references} {.qa} {.backup}.
  - bullets            → bullet slide (keep ≤ 6 bullets, ≤ 40 words; details go in notes)
  - ![caption](figures/x.png)  → picture fitted without cropping (report/figures/ is searched too)
  - | a | b | tables   → native, editable table
  - text, then |||, then more → two columns
  - .results: "- 0.92 :: FedProx accuracy" → big-number cards
  - [@key] cites refs.bib (slides/ or ../report/); the {.references} slide lists what was cited
  - ::: notes ... ::: → speaker notes
Designs: {.cards} "- **Label**: text" · {.steps} · {.timeline} "- Jan :: text" · {.statement} ·
{.quote} · {.image-left} {.image-right}. Own background: {.dark} {.accent} {.light} {color="#E63946"}
{bg="photo.jpg"}. Plain slides that fit a design are drawn that way (variety: auto); {.bullets} keeps bullets.
-->

# Introduction

## [[TODO]] One-sentence problem statement as the title

- Why the problem matters
- What is missing in current approaches
- What this work does

::: notes
What you will say on this slide.
:::

## Objectives

1. First objective
2. Second objective

# Results

## [[TODO]] Headline result as the title {.results}

- 0.00 :: Main metric
- 0.00 :: Baseline

## References {.references}
"""


def deref(text):
    """Report cross-references read badly on slides: '@fig-acc shows' → 'The figure shows'."""
    text = re.sub(r"@fig-[\w-]+", "the figure", text)
    text = re.sub(r"@tbl-[\w-]+", "the table", text)
    text = re.sub(r"@eq-[\w-]+", "the equation", text)
    text = re.sub(r"@(?:alg|sec)-[\w-]+", "the algorithm", text)
    return re.sub(r"(^|[.!?]\s+)([a-z])", lambda m: m.group(1) + m.group(2).upper(), text)


def first_sentence(par):
    m = re.match(r"(.+?[.!?])(\s+[A-Z(]|$)", par)
    return (m.group(1) if m else par).strip()


def section_parts(lines):
    """A report section's lines → paragraphs, list items, equations, pipe tables, figures."""
    parts = {"paras": [], "items": [], "eqs": [], "tables": [], "figs": [], "skipped": []}
    buf, eq, div, table = [], None, False, []

    def flush():
        if buf:
            parts["paras"].append(deref(" ".join(buf)))
            buf.clear()

    for line in lines + [""]:
        st = line.strip()
        if eq is not None:                                   # inside a multi-line $$ … $$
            if st.startswith("$$"):
                parts["eqs"].append(" ".join(eq).strip())
                eq = None
            else:
                eq.append(st)
            continue
        if div:                                              # ::: {#alg-…} … ::: is skipped
            div = not st.startswith(":::") or st.startswith("::: {")
            continue
        if st.startswith("::: {") or st.startswith(":::{"):
            flush()
            div = True
            parts["skipped"].append("an algorithm")
            continue
        if st.startswith("$$"):
            flush()
            body = re.sub(r"\s*\{#eq-[^}]*\}\s*$", "", st)
            if len(body) > 4 and body.endswith("$$"):
                parts["eqs"].append(body[2:-2].strip())
            else:
                eq = [body[2:]]
            continue
        if st.startswith("|") or st.startswith("+"):
            flush()
            table.append(st)
            continue
        if table:                                            # a table just ended
            if any(t.startswith("+") for t in table):
                parts["skipped"].append("a grid table (merged cells): retype it as a | pipe | table")
            elif any(re.fullmatch(r"\|?\s*:?-{2,}.*", t) for t in table):
                parts["tables"].append(table)
            table = []
        img = re.match(r"^!\[(.*?)\]\((.*?)\)", st)
        if img:
            flush()
            parts["figs"].append((img.group(1), img.group(2)))
        elif re.match(r"^(\d+[.)]|[-*+])\s", st):
            flush()
            parts["items"].append(deref(re.sub(r"^(\d+[.)]|[-*+])\s+", "", st)))
        elif st.startswith(":"):                               # table caption line
            continue
        elif st:
            buf.append(st)
        else:
            flush()
    return parts


def draft_from_report(rdir: Path):
    """Report chapters → a slides.md draft: one slide per section (bullets = list items and each
    paragraph's first sentence, notes = the full paragraphs), a picture slide per figure. Titles are
    left for a person (or Claude) to turn into claims, so they carry [[TODO]] and the build flags them."""
    cfg = yaml.safe_load((rdir / "report.yml").read_text()) or {}
    order = cfg.get("chapters")
    files = [rdir / "chapters" / f for f in order] if order else sorted((rdir / "chapters").glob("*.md"))
    out = []
    for f in files:
        text = re.sub(r"<!--.*?-->", "", f.read_text(), flags=re.S)
        sections, cur = [], None
        for line in text.splitlines():
            m = re.match(r"^(#{1,2})\s+(.*?)\s*(\{[^}]*\})?\s*$", line)
            if m:
                cur = {"level": len(m.group(1)), "title": m.group(2), "lines": []}
                sections.append(cur)
            elif cur is not None:
                cur["lines"].append(line)
        for sec in sections:
            if sec["level"] == 1:
                out.append(f"# {sec['title']}\n")
            p = section_parts(sec["lines"])
            bullets = (p["items"] + [first_sentence(x) for x in p["paras"]])[:5]
            if bullets or p["eqs"] or p["tables"] or p["skipped"]:
                out.append(f"## [[TODO: state the point]] {sec['title']}\n")
                out += [f"- {b}" for b in bullets] + ([""] if bullets else [])
                for e in p["eqs"]:
                    out += [f"$${e}$$", ""]
                for t in p["tables"][:1]:
                    out += t + [""]
                for sk in p["skipped"]:
                    out += [f"<!-- [[TODO]] the report has {sk} here -->", ""]
                if p["paras"] or p["items"]:
                    out += ["::: notes", *p["paras"], *[f"- {i}" for i in p["items"]], ":::", ""]
            for cap, path in p["figs"]:
                out += [f"## [[TODO: what this shows]] {cap.rstrip('.')}\n", f"![{cap}]({path})", ""]
    out.append("## References {.references}\n")
    return "\n".join(out)


def init(target: Path, from_report=False):
    if (target / "slides.yml").exists():
        sys.exit(f"{target}/slides.yml already exists")
    for sub in ("figures", "diagrams"):
        (target / sub).mkdir(parents=True, exist_ok=True)
    title = "Untitled talk"
    rep = target.parent / "report" / "report.yml"
    if rep.exists():
        title = (yaml.safe_load(rep.read_text()) or {}).get("title", title)
    (target / "slides.yml").write_text(STARTER_YML.format(title=title))
    if from_report:
        if not rep.exists():
            sys.exit("--from-report: no report/report.yml next to the slides folder")
        (target / "slides.md").write_text(draft_from_report(rep.parent))
        (target / "assumptions.md").write_text(
            "# Assumptions (delete a line once you have checked it)\n"
            "- Drafted from the report: one slide per section, bullets are each paragraph's first sentence,\n"
            "  speaker notes hold the full paragraphs. Cut to the 3-4 points you will actually say.\n"
            "- Every title is still a section name; rewrite each as the slide's claim.\n")
    else:
        (target / "slides.md").write_text(STARTER_MD)
    (target / ".gitignore").write_text("out/\n_build/\nfigures/.*.done\n")
    say(f"created {target}/: write slides.md, then run `report slides build`")


# ── theme gallery: report slides themes ─────────────────────
GALLERY_MD = """\
# Problem

## Raw patient data cannot leave the hospital

Hospitals hold the data a good model needs, but privacy rules keep every record on site.

- **Privacy**: records cannot be pooled in one place
- **Data**: each hospital sees only its own patients
- **Accuracy**: a model trained alone generalises poorly

# Method

## FedProx keeps clients together where FedAvg lets them drift

### FedAvg (existing)

- Clients train freely between rounds
- Models drift apart on non-IID data

### FedProx (proposed)

- A proximal term limits each client's drift
- Accuracy rises steadily

# Results

## Conclusion

- **Stable**: FedProx converges where FedAvg oscillates
- **Accurate**: 3 points higher after 20 rounds
- **Cheap**: one extra term, no extra communication
"""

GALLERY_ROWS = [("clean", "ocean"), ("band", "teal"), ("sidebar", "berry"), ("minimal", "crimson"),
                ("bold", "royal"), ("soft", "mint"), ("geometric", "sunset")]


def themes_gallery(dest: Path):
    """Every style (each in a different palette) as a row of sample slides → dest/themes.png, plus the
    palettes → dest/palettes.png."""
    import deck_theme
    import pymupdf
    from PIL import ImageDraw
    dest.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        decks = []
        for style, pal in GALLERY_ROWS:
            sdir = Path(tmp) / f"{style}-{pal}"
            sdir.mkdir()
            (sdir / "slides.yml").write_text(yaml.safe_dump({"title": "Federated Learning on Non-IID Data",
                                                             "subtitle": "Project review", "qa": False,
                                                             "agenda": False,
                                                             "theme": {"style": style, "palette": pal}}))
            (sdir / "slides.md").write_text(GALLERY_MD)
            cfg = load_config(sdir)
            cfg["_report"] = None
            problems = []
            b = assemble(sdir, cfg, parse(GALLERY_MD), problems)
            pptx = sdir / f"{style}-{pal}.pptx"
            b.prs.save(pptx)
            decks.append((style, pal, pptx))
        say("rendering the samples with LibreOffice")
        with tempfile.TemporaryDirectory() as prof:
            subprocess.run(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", "pdf",
                            "--outdir", tmp, *[str(p) for *_, p in decks]], capture_output=True, timeout=600)
        rows = []
        for style, pal, pptx in decks:
            pdf = Path(tmp) / (pptx.stem + ".pdf")
            ims = []
            for p in pymupdf.open(pdf):
                pix = p.get_pixmap(dpi=30)
                ims.append(Image.frombytes("RGB", (pix.width, pix.height), pix.samples))
            rows.append((f"{style}  ·  {pal}", ims))
    w, h = rows[0][1][0].size
    pad, label = 8, 34
    cols = max(len(ims) for _, ims in rows)
    sheet = Image.new("RGB", (cols * (w + pad) + pad, len(rows) * (h + pad + label) + pad), "white")
    d = ImageDraw.Draw(sheet)
    from PIL import ImageFont
    f = subprocess.run(["fc-match", "-f", "%{file}", "Carlito:style=Bold"], capture_output=True, text=True).stdout
    font = ImageFont.truetype(f, 24) if f else ImageFont.load_default()
    for r, (name, ims) in enumerate(rows):
        y = pad + r * (h + pad + label)
        d.text((pad, y + 4), name, fill="#222", font=font)
        for c, im in enumerate(ims):
            sheet.paste(im, (pad + c * (w + pad), y + label))
    sheet.save(dest / "themes.png")
    deck_theme.swatches(dest / "palettes.png")
    say(f"wrote {dest / 'themes.png'} and {dest / 'palettes.png'}")
