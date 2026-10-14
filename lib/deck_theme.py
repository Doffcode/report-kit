"""Kit themes: a slide theme made at build time from a style, a palette and a font pair.

slides.yml:
    theme:
      style: band          # clean | band | sidebar | minimal | bold | soft | geometric
      palette: teal        # one of PALETTES, or your own colours:
      # primary: "#7B2CBF" #   one colour is enough, the others are derived from it
      # accent: "#F72585"  #   (also: dark, light)
      fonts: classic       # classic | serif-titles | formal | plain | times

or short: `theme: band teal`. `report slides themes` draws every style and palette.

python-pptx can't create masters or layouts, so each theme starts from python-pptx's
default Office template (standard layout names) and rewrites it: 16:9, colours,
fonts, text styles, background shapes and the title/section slide designs.
Background shapes are named "Kit Bg …" so the builder knows text may sit on them.
"""
import colorsys
import copy
import hashlib
import json
import re
import sys
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Emu, Inches

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main"}

STYLES = {
    "clean": "white slides, title with a short accent bar; dark title and section slides",
    "band": "title in a dark band across the top",
    "sidebar": "coloured strip down the left; split title slide",
    "minimal": "no decoration, large light titles, thin rule above the footer",
    "bold": "dark slides with light text everywhere",
    "soft": "tinted slides with soft circles in the corner",
    "geometric": "triangles in the corner; angled title slide",
}
PALETTES = {                    # dark (titles, dark slides), primary, accent, light (tints, cards)
    "ocean": ("1F3A5F", "2E6DB4", "E07A1F", "EEF2F7"),
    "royal": ("14213D", "1D4ED8", "FCA311", "EEF2FB"),
    "teal": ("0B3C49", "0F8B8D", "EC9A29", "E8F4F4"),
    "forest": ("1E3D2F", "2D6A4F", "D4A373", "EEF4EF"),
    "mint": ("1B4332", "40916C", "F4845F", "EDF7F1"),
    "sunset": ("3D1F2B", "C8553D", "F2A541", "FBF0EA"),
    "crimson": ("2B0F14", "A4161A", "3D5A80", "F7ECEC"),
    "berry": ("2B1B3D", "7B2CBF", "F72585", "F3EEF8"),
    "plum": ("3C1642", "8E3B8A", "2EC4B6", "F5EDF5"),
    "slate": ("22303C", "4A6274", "E2A03F", "EEF1F4"),
    "earth": ("3E2C23", "8C5E3C", "5B8C5A", "F6F0EA"),
    "gold": ("0A1F44", "13315C", "C9A227", "F3F1EA"),
    "mono": ("1E1E1E", "404040", "D62828", "F2F2F2"),
}
FONT_PAIRS = {                  # (titles, body): fonts every college PC has, with same-width Linux twins
    "classic": ("Calibri", "Calibri"),
    "serif-titles": ("Cambria", "Calibri"),
    "formal": ("Cambria", "Cambria"),
    "plain": ("Arial", "Arial"),
    "times": ("Times New Roman", "Times New Roman"),
}
DEFAULT = {"style": "clean", "palette": "ocean", "fonts": "classic"}

KEEP = ["Title Slide", "Title and Content", "Section Header", "Two Content", "Title Only", "Blank"]
W, H = Emu(12192000), Emu(6858000)   # 13.333 x 7.5 in, exact 16:9
M = Inches(0.6)
GAP = Inches(0.4)
IN = 914400


# ── colours ──────────────────────────────────────────────────
def rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def hexc(c):
    return "%02X%02X%02X" % tuple(max(0, min(255, round(x))) for x in c)


def mix(a, b, t):
    """t = 0 → a, t = 1 → b."""
    return hexc([x + (y - x) * t for x, y in zip(rgb(a), rgb(b))])


def luminance(h):
    def ch(v):
        v /= 255
        return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4
    r, g, b = rgb(h)
    return 0.2126 * ch(r) + 0.7152 * ch(g) + 0.0722 * ch(b)


def hls(h, dh=0, light=None, sat=None):
    r, g, b = (x / 255 for x in rgb(h))
    hh, ll, ss = colorsys.rgb_to_hls(r, g, b)
    hh = (hh + dh / 360) % 1
    ll = ll if light is None else light
    ss = ss if sat is None else sat
    return hexc([x * 255 for x in colorsys.hls_to_rgb(hh, ll, ss)])


def palette(spec):
    """→ {dark, primary, accent, light} from a named palette and/or explicit colours."""
    name = spec.get("palette", DEFAULT["palette"])
    if name not in PALETTES and not spec.get("primary"):
        sys.exit(f"theme palette {name!r}: unknown. Pick one of {', '.join(PALETTES)} or give primary: \"#hex\"")
    base = dict(zip(("dark", "primary", "accent", "light"), PALETTES.get(name, PALETTES["ocean"])))
    if spec.get("primary"):                                   # own colour: derive what isn't given
        p = str(spec["primary"]).lstrip("#").upper()
        _, l, s = colorsys.rgb_to_hls(*(x / 255 for x in rgb(p)))
        base = {"primary": p, "dark": hls(p, light=min(l, 0.17), sat=min(s, 0.55)),
                "accent": hls(p, dh=150, light=0.55, sat=max(min(s, 0.8), 0.55)), "light": mix(p, "FFFFFF", 0.92)}
    for k in ("dark", "accent", "light"):
        if spec.get(k):
            base[k] = str(spec[k]).lstrip("#").upper()
    return base


def scheme(pal):
    p = pal["primary"]
    return {"dk1": mix(pal["dark"], "000000", 0.55), "lt1": "FFFFFF", "dk2": pal["dark"], "lt2": pal["light"],
            "accent1": p, "accent2": pal["accent"], "accent3": hls(p, dh=110, light=0.42, sat=0.5),
            "accent4": hls(p, dh=-70, light=0.5, sat=0.55), "accent5": mix(pal["dark"], "FFFFFF", 0.45),
            "accent6": "7F8C8D", "hlink": p, "folHlink": hls(p, dh=40)}


# ── spec from slides.yml ─────────────────────────────────────
def spec_of(cfg):
    t = cfg.get("theme") or {}
    if isinstance(t, str):                                    # "band teal" / "band" / "teal"
        words, t = t.split(), {}
        for w in words:
            if w in STYLES:
                t["style"] = w
            elif w in PALETTES:
                t["palette"] = w
            elif w in FONT_PAIRS:
                t["fonts"] = w
            elif re.fullmatch(r"#?[0-9A-Fa-f]{6}", w):
                t["primary"] = w
            else:
                sys.exit(f"theme: {w!r} is not a style ({', '.join(STYLES)}), palette ({', '.join(PALETTES)}) "
                         f"or font pair ({', '.join(FONT_PAIRS)})")
    spec = {**DEFAULT, **{k: v for k, v in t.items() if v is not None}}
    if spec["style"] not in STYLES:
        sys.exit(f"theme style {spec['style']!r}: unknown. Pick one of {', '.join(STYLES)}")
    if spec["fonts"] not in FONT_PAIRS:
        sys.exit(f"theme fonts {spec['fonts']!r}: unknown. Pick one of {', '.join(FONT_PAIRS)}")
    return spec


def make(spec, out_dir: Path) -> Path:
    """Write the theme .pptx (cached by spec + this file's code) and return its path."""
    key = hashlib.sha1((json.dumps(spec, sort_keys=True) + Path(__file__).read_text()).encode()).hexdigest()[:10]
    name = spec.get("palette") if not spec.get("primary") else "custom"
    out = out_dir / f"theme-{spec['style']}-{name}-{key}.pptx"
    if out.exists():
        return out
    out_dir.mkdir(parents=True, exist_ok=True)
    for old in out_dir.glob("theme-*.pptx"):
        old.unlink()
    pal = palette(spec)
    T = Theme(pal, FONT_PAIRS[spec["fonts"]], spec["style"])
    T.build()
    T.prs.save(out)
    return out


# ── XML helpers ──────────────────────────────────────────────
def lvl_style(parent, lvl=1, sz=None, colour=None, bold=None, algn=None, italic=None):
    """Set a:lvlNpPr inside a p:titleStyle/p:bodyStyle or a shape's a:lstStyle."""
    tag = f"{A}lvl{lvl}pPr"
    ppr = parent.find(tag)
    if ppr is None:
        ppr = etree.SubElement(parent, tag)
    if algn:
        ppr.set("algn", algn)
    rpr = ppr.find(f"{A}defRPr")
    if rpr is None:
        rpr = etree.SubElement(ppr, f"{A}defRPr")
    if sz:
        rpr.set("sz", str(int(sz * 100)))
    rpr.attrib.pop("cap", None)                      # the Office section header shouts in ALL CAPS
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
    if italic is not None:
        rpr.set("i", "1" if italic else "0")
    if colour:
        for f in rpr.findall(f"{A}solidFill"):
            rpr.remove(f)
        fill = etree.Element(f"{A}solidFill")
        hexv = re.fullmatch(r"[0-9A-F]{6}", colour)
        etree.SubElement(fill, f"{A}srgbClr" if hexv else f"{A}schemeClr").set("val", colour)
        rpr.insert(0, fill)
    return ppr


def ph_style(ph, anchor=None, **kw):
    txbody = ph._element.txBody
    lst = txbody.find(f"{A}lstStyle")
    if lst is None:
        lst = etree.Element(f"{A}lstStyle")
        txbody.insert(1, lst)
    lvl_style(lst, **kw)
    if anchor:
        txbody.find(f"{A}bodyPr").set("anchor", anchor)


def no_autofit(ph):
    bp = ph._element.txBody.find(f"{A}bodyPr")
    for t in ("normAutofit", "spAutoFit", "noAutofit"):
        for e in bp.findall(f"{A}{t}"):
            bp.remove(e)
    etree.SubElement(bp, f"{A}noAutofit")


def place(shape, box):
    shape.left, shape.top, shape.width, shape.height = (Emu(int(v)) for v in box)


def is_footer(ph):
    return any(k in str(ph.placeholder_format.type) for k in ("SLIDE_NUMBER", "FOOTER", "DATE"))


def footer_boxes(container, colour):
    for ph in container.placeholders:
        t = str(ph.placeholder_format.type)
        if "SLIDE_NUMBER" in t:
            place(ph, (W - M - Inches(1.2), H - Inches(0.5), Inches(1.2), Inches(0.35)))
            ph_style(ph, sz=12, colour=colour, algn="r")
        elif "FOOTER" in t:
            place(ph, (M, H - Inches(0.5), Inches(8), Inches(0.35)))
            ph_style(ph, sz=12, colour=colour, algn="l")
        elif "DATE" in t:
            place(ph, (W - M - Inches(4.4), H - Inches(0.5), Inches(3), Inches(0.35)))
            ph_style(ph, sz=12, colour=colour)


class Theme:
    def __init__(self, pal, fonts, style):
        self.pal, self.fonts, self.style = pal, fonts, style
        self.c = scheme(pal)
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self._tmp = None

    # shapes can only be drawn on slides: draw on a scratch slide, move the XML to the master/layout
    def shape(self, target, kind, box, colour, alpha=None, name="Kit Bg", flip="", rot=0):
        if self._tmp is None:
            self._tmp = self.prs.slides.add_slide(self.prs.slide_layouts[-1])
        sh = self._tmp.shapes.add_shape(kind, *(Emu(int(v)) for v in box))
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(colour)
        sh.line.fill.background()
        sh.name = name
        if rot:
            sh.rotation = rot
        el = sh._element
        style = el.find("{%s}style" % NS["p"])
        if style is not None:                          # drop the Office shape style (outline, shadow)
            el.remove(style)
        if alpha is not None:
            el.find(".//" + A + "srgbClr").append(etree.Element(f"{A}alpha", val=str(int(alpha * 100000))))
        xfrm = el.find(".//" + A + "xfrm")
        if "h" in flip:
            xfrm.set("flipH", "1")
        if "v" in flip:
            xfrm.set("flipV", "1")
        if kind == MSO_SHAPE.ROUNDED_RECTANGLE:
            sh.adjustments[0] = 0.5
        el.getparent().remove(el)
        tree = target.shapes._spTree
        ids = [int(e.get("id")) for e in tree.iter("{%s}cNvPr" % NS["p"])]
        el.find(".//{%s}cNvPr" % NS["p"]).set("id", str(max(ids + [1]) + 1 + len(ids)))
        # behind the placeholders: right after nvGrpSpPr/grpSpPr, in drawing order
        n_bg = sum(1 for e in tree if e.find(".//{%s}cNvPr" % NS["p"]) is not None
                   and (e.find(".//{%s}cNvPr" % NS["p"]).get("name") or "").startswith("Kit Bg"))
        tree.insert(2 + n_bg, el)
        return el

    def background(self, target, colour):
        target.background.fill.solid()
        target.background.fill.fore_color.rgb = RGBColor.from_string(colour)

    def build(self):
        self.set_theme()
        self.text_styles()
        layouts = self.keep_layouts()
        getattr(self, "style_" + self.style)(layouts)
        if self._tmp is not None:
            rid = self.prs.slides._sldIdLst[-1].rId
            self.prs.part.drop_rel(rid)
            del self.prs.slides._sldIdLst[-1]

    def set_theme(self):
        part = [r for r in self.prs.slide_master.part.rels.values() if r.reltype.endswith("/theme")][0].target_part
        xml = part.blob.decode()
        for name, hexv in self.c.items():
            xml = re.sub(rf"<a:{name}>.*?</a:{name}>", f'<a:{name}><a:srgbClr val="{hexv}"/></a:{name}>', xml,
                         flags=re.S)
        major, minor = self.fonts
        xml = re.sub(r'(<a:majorFont>\s*<a:latin typeface=")[^"]*', rf"\g<1>{major}", xml)
        xml = re.sub(r'(<a:minorFont>\s*<a:latin typeface=")[^"]*', rf"\g<1>{minor}", xml)
        xml = re.sub(r'name="Office Theme"', 'name="Report Kit"', xml)
        part._blob = xml.encode()

    # geometry of content slides; styles change it before text_styles() runs
    title_box = (M, Inches(0.35), W - 2 * M, Inches(1.0))
    title_anchor = "b"
    body_top = Inches(1.6)

    def text_styles(self, title_sz=32, title_colour=None, title_bold=True, body_colour=None, bullet="accent1"):
        m = self.prs.slide_master
        tx = m._element.find(".//p:txStyles", NS)
        lvl_style(tx.find("p:titleStyle", NS), sz=title_sz, colour=title_colour or "tx2", bold=title_bold, algn="l")
        body = tx.find("p:bodyStyle", NS)
        for lvl, sz in ((1, 24), (2, 20), (3, 18), (4, 16), (5, 16)):
            ppr = lvl_style(body, lvl=lvl, sz=sz, colour=body_colour)
            for tag in ("spcBef", "buClr"):
                for e in ppr.findall(f"{A}{tag}"):
                    ppr.remove(e)
            spc = etree.Element(f"{A}spcBef")
            etree.SubElement(spc, f"{A}spcPts").set("val", "1000" if lvl == 1 else "400")
            ppr.insert(0, spc)
            buclr = etree.Element(f"{A}buClr")
            hexv = re.fullmatch(r"[0-9A-F]{6}", bullet)
            etree.SubElement(buclr, f"{A}srgbClr" if hexv else f"{A}schemeClr").set("val", bullet)
            ppr.insert(1, buclr)

    def content_geometry(self, layouts, footer_colour="7F8C8D"):
        """Title and body boxes on the master and the content layouts; autofit off everywhere."""
        m = self.prs.slide_master
        body = (M, self.body_top, W - 2 * M, H - Inches(0.75) - self.body_top)
        for ph in m.placeholders:
            t = str(ph.placeholder_format.type)
            if "TITLE" in t:
                place(ph, self.title_box)
                ph._element.txBody.find(f"{A}bodyPr").set("anchor", self.title_anchor)
            elif "BODY" in t:
                place(ph, body)
            if ph.has_text_frame:
                no_autofit(ph)
        footer_boxes(m, footer_colour)
        for name in ("Title and Content", "Title Only", "Two Content", "Blank"):
            lay = layouts[name]
            footer_boxes(lay, footer_colour)
            for ph in lay.placeholders:
                if ph.has_text_frame:
                    no_autofit(ph)
            phs = {ph.placeholder_format.idx: ph for ph in lay.placeholders}
            if 0 in phs:
                place(phs[0], self.title_box)
                phs[0]._element.txBody.find(f"{A}bodyPr").set("anchor", self.title_anchor)
            if name == "Title and Content":
                place(phs[1], body)
            elif name == "Two Content":
                col = (body[2] - GAP) / 2
                place(phs[1], (body[0], body[1], col, body[3]))
                place(phs[2], (body[0] + col + GAP, body[1], col, body[3]))

    def keep_layouts(self):
        lays = self.prs.slide_layouts
        for lay in list(lays):
            if lay.name not in KEEP:
                lays.remove(lay)
        for lay in lays:
            for ph in lay.placeholders:
                if ph.has_text_frame:
                    no_autofit(ph)
        return {lay.name: lay for lay in lays}

    def cover(self, lay, bg, title, sub, footer_colour, hide_master=True):
        """Title Slide / Section Header: background, title box (idx 0) and subtitle/label box (idx 1)."""
        if bg:
            self.background(lay, bg)
        if hide_master:
            lay._element.set("showMasterSp", "0")
        phs = {ph.placeholder_format.idx: ph for ph in lay.placeholders}
        for idx, (box, kw) in ((0, title), (1, sub)):
            place(phs[idx], box)
            ph_style(phs[idx], **kw)
        footer_boxes(lay, footer_colour)

    def bar(self, target, x, y, w=Inches(1.1), h=Inches(0.06), colour=None):
        self.shape(target, MSO_SHAPE.RECTANGLE, (x, y, w, h), colour or self.pal["accent"], name="Kit Bg Bar")

    # ── styles ───────────────────────────────────────────────
    def style_clean(self, L):
        P = self.pal
        self.content_geometry(L)
        self.bar(self.prs.slide_master, M, Inches(1.42))
        sub_c = mix(P["light"], P["dark"], 0.12)
        self.cover(L["Title Slide"], P["dark"],
                   ((Inches(0.9), Inches(1.7), W - Inches(1.8) - Inches(1.6), Inches(1.9)),
                    dict(sz=40, colour="FFFFFF", bold=True, algn="l", anchor="b")),
                   ((Inches(0.9), Inches(3.95), W - Inches(1.8), Inches(2.6)),
                    dict(sz=20, colour=sub_c, algn="l", anchor="t")), mix(P["light"], P["dark"], 0.35))
        self.bar(L["Title Slide"], Inches(0.9), Inches(3.72), Inches(1.4), Inches(0.07))
        self.cover(L["Section Header"], P["dark"],
                   ((Inches(0.9), Inches(2.9), W - Inches(1.8), Inches(1.8)),
                    dict(sz=40, colour="FFFFFF", bold=True, algn="l", anchor="t")),
                   ((Inches(0.9), Inches(2.2), W - Inches(1.8), Inches(0.7)),
                    dict(sz=20, colour=mix(P["accent"], "FFFFFF", 0.35), algn="l", anchor="b")),
                   mix(P["light"], P["dark"], 0.35))

    def style_band(self, L):
        P = self.pal
        band = Inches(1.25)
        self.title_box = (M, Inches(0.12), W - 2 * M, Inches(1.0))
        self.title_anchor = "ctr"
        self.body_top = Inches(1.65)
        self.text_styles(title_sz=30, title_colour="FFFFFF")
        self.content_geometry(L)
        m = self.prs.slide_master
        self.shape(m, MSO_SHAPE.RECTANGLE, (0, 0, W, band), P["dark"], name="Kit Bg Band")
        self.shape(m, MSO_SHAPE.RECTANGLE, (0, band, W, Inches(0.05)), P["accent"], name="Kit Bg Line")
        t = L["Title Slide"]
        self.cover(t, "FFFFFF",
                   ((Inches(0.9), Inches(1.0), W - Inches(1.8) - Inches(1.8), Inches(2.9)),
                    dict(sz=40, colour="FFFFFF", bold=True, algn="l", anchor="b")),
                   ((Inches(0.9), Inches(4.6), W - Inches(1.8), Inches(2.4)),
                    dict(sz=20, colour=self.c["dk1"], algn="l", anchor="t")), "7F8C8D")
        self.shape(t, MSO_SHAPE.RECTANGLE, (0, 0, W, Inches(4.2)), P["dark"], name="Kit Bg Band")
        self.shape(t, MSO_SHAPE.RECTANGLE, (0, Inches(4.2), W, Inches(0.08)), P["accent"], name="Kit Bg Line")
        s = L["Section Header"]
        self.cover(s, P["primary"],
                   ((Inches(0.9), Inches(3.0), W - Inches(1.8), Inches(1.8)),
                    dict(sz=40, colour="FFFFFF", bold=True, algn="l", anchor="t")),
                   ((Inches(0.9), Inches(2.1), W - Inches(1.8), Inches(0.7)),
                    dict(sz=20, colour=mix(P["primary"], "FFFFFF", 0.75), algn="l", anchor="b")),
                   mix(P["primary"], "FFFFFF", 0.6))
        self.bar(s, Inches(0.9), Inches(2.88), Inches(1.4), Inches(0.07))

    def style_sidebar(self, L):
        P = self.pal
        self.content_geometry(L)
        m = self.prs.slide_master
        self.shape(m, MSO_SHAPE.RECTANGLE, (0, 0, Inches(0.3), H), P["primary"], name="Kit Bg Strip")
        self.shape(m, MSO_SHAPE.RECTANGLE, (0, Inches(0.5), Inches(0.3), Inches(0.9)), P["accent"],
                   name="Kit Bg Mark")
        t = L["Title Slide"]
        panel = int(W * 0.42)
        self.cover(t, "FFFFFF",
                   ((Inches(0.7), Inches(0.9), panel - Inches(1.2), Inches(3.9)),
                    dict(sz=36, colour="FFFFFF", bold=True, algn="l", anchor="b")),
                   ((panel + Inches(0.7), Inches(2.1), W - panel - Inches(1.3), Inches(4.4)),
                    dict(sz=20, colour=self.c["dk1"], algn="l", anchor="t")), "7F8C8D")
        self.shape(t, MSO_SHAPE.RECTANGLE, (0, 0, panel, H), P["dark"], name="Kit Bg Panel")
        self.shape(t, MSO_SHAPE.RECTANGLE, (panel, 0, Inches(0.12), H), P["accent"], name="Kit Bg Edge")
        self.bar(t, Inches(0.7), Inches(5.0), Inches(1.2), Inches(0.07))
        s = L["Section Header"]
        self.cover(s, P["light"],
                   ((Inches(1.9), Inches(2.9), W - Inches(2.6), Inches(1.8)),
                    dict(sz=40, colour=P["dark"], bold=True, algn="l", anchor="t")),
                   ((Inches(1.9), Inches(2.2), W - Inches(2.6), Inches(0.7)),
                    dict(sz=20, colour=P["primary"], bold=True, algn="l", anchor="b")), "7F8C8D")
        self.shape(s, MSO_SHAPE.RECTANGLE, (0, 0, Inches(1.2), H), P["primary"], name="Kit Bg Strip")
        for ph in s.placeholders:
            if "FOOTER" in str(ph.placeholder_format.type):
                place(ph, (Inches(1.9), H - Inches(0.5), Inches(8), Inches(0.35)))
        self.shape(s, MSO_SHAPE.RECTANGLE, (Inches(1.2), 0, Inches(0.1), H), P["accent"], name="Kit Bg Edge")

    def style_minimal(self, L):
        P = self.pal
        self.text_styles(title_sz=34, title_colour=P["dark"], title_bold=False)
        self.content_geometry(L, footer_colour="8C8C8C")
        self.shape(self.prs.slide_master, MSO_SHAPE.RECTANGLE,
                   (M, H - Inches(0.6), W - 2 * M, Emu(12700)), "BFBFBF", name="Kit Bg Rule")
        t = L["Title Slide"]
        self.cover(t, "FFFFFF",
                   ((Inches(0.9), Inches(1.3), W - Inches(1.8) - Inches(1.6), Inches(2.6)),
                    dict(sz=44, colour=P["dark"], bold=True, algn="l", anchor="b")),
                   ((Inches(0.9), Inches(4.35), W - Inches(1.8), Inches(2.6)),
                    dict(sz=20, colour="595959", algn="l", anchor="t")), "8C8C8C")
        self.bar(t, Inches(0.9), Inches(4.1), Inches(0.9), Inches(0.1))
        s = L["Section Header"]
        self.cover(s, "FFFFFF",
                   ((Inches(0.9), Inches(2.9), W - Inches(1.8), Inches(1.8)),
                    dict(sz=44, colour=P["primary"], bold=True, algn="l", anchor="t")),
                   ((Inches(0.9), Inches(2.2), W - Inches(1.8), Inches(0.7)),
                    dict(sz=22, colour=P["accent"], bold=True, algn="l", anchor="b")), "8C8C8C")
        self.shape(s, MSO_SHAPE.RECTANGLE, (Inches(0.9), Inches(4.8), W * 0.42, Emu(12700)), "BFBFBF",
                   name="Kit Bg Rule")

    def style_bold(self, L):
        P = self.pal
        light_text = mix(P["light"], "FFFFFF", 0.4)
        self.text_styles(title_colour="FFFFFF", body_colour=light_text, bullet=P["accent"])
        self.content_geometry(L, footer_colour=mix(P["dark"], "FFFFFF", 0.5))
        self.background(self.prs.slide_master, P["dark"])
        self.bar(self.prs.slide_master, M, Inches(1.42))
        deep = mix(P["dark"], "000000", 0.35)
        t = L["Title Slide"]
        self.cover(t, deep,
                   ((Inches(1.25), Inches(1.6), W - Inches(2.2) - Inches(1.6), Inches(2.1)),
                    dict(sz=44, colour="FFFFFF", bold=True, algn="l", anchor="b")),
                   ((Inches(1.25), Inches(4.0), W - Inches(2.2), Inches(2.6)),
                    dict(sz=20, colour=mix(P["light"], deep, 0.2), algn="l", anchor="t")),
                   mix(deep, "FFFFFF", 0.45))
        self.shape(t, MSO_SHAPE.RECTANGLE, (Inches(0.9), Inches(1.75), Inches(0.12), Inches(1.95)), P["accent"],
                   name="Kit Bg Bar")
        s = L["Section Header"]
        self.cover(s, P["primary"],
                   ((Inches(0.9), Inches(2.9), W - Inches(1.8), Inches(1.8)),
                    dict(sz=44, colour="FFFFFF", bold=True, algn="l", anchor="t")),
                   ((Inches(0.9), Inches(2.2), W - Inches(1.8), Inches(0.7)),
                    dict(sz=20, colour=mix(P["primary"], "FFFFFF", 0.75), algn="l", anchor="b")),
                   mix(P["primary"], "FFFFFF", 0.6))

    def style_soft(self, L):
        P = self.pal
        self.title_box = (M, Inches(0.35), W - 2 * M - Inches(1.6), Inches(1.0))
        self.content_geometry(L)
        m = self.prs.slide_master
        self.background(m, P["light"])
        self.shape(m, MSO_SHAPE.OVAL, (W - Inches(1.9), -Inches(1.3), Inches(3.0), Inches(3.0)), P["primary"],
                   alpha=0.22, name="Kit Bg Circle")
        self.shape(m, MSO_SHAPE.OVAL, (W - Inches(2.15), Inches(0.95), Inches(0.45), Inches(0.45)), P["accent"],
                   alpha=0.7, name="Kit Bg Dot")
        self.shape(m, MSO_SHAPE.ROUNDED_RECTANGLE, (M, Inches(1.42), Inches(1.1), Inches(0.08)), P["accent"],
                   name="Kit Bg Bar")
        t = L["Title Slide"]
        self.cover(t, P["primary"],
                   ((Inches(0.9), Inches(1.6), W - Inches(1.8) - Inches(1.6), Inches(2.1)),
                    dict(sz=42, colour="FFFFFF", bold=True, algn="l", anchor="b")),
                   ((Inches(0.9), Inches(4.0), W - Inches(1.8), Inches(2.6)),
                    dict(sz=20, colour=mix(P["primary"], "FFFFFF", 0.85), algn="l", anchor="t")),
                   mix(P["primary"], "FFFFFF", 0.6))
        self.shape(t, MSO_SHAPE.OVAL, (W - Inches(5.2), H - Inches(3.6), Inches(7), Inches(7)), "FFFFFF",
                   alpha=0.10, name="Kit Bg Circle")
        self.shape(t, MSO_SHAPE.OVAL, (-Inches(1.6), -Inches(1.9), Inches(3.8), Inches(3.8)), "FFFFFF",
                   alpha=0.10, name="Kit Bg Circle")
        self.shape(t, MSO_SHAPE.ROUNDED_RECTANGLE, (Inches(0.9), Inches(3.75), Inches(1.4), Inches(0.09)),
                   P["accent"], name="Kit Bg Bar")
        s = L["Section Header"]
        self.cover(s, P["dark"],
                   ((Inches(0.9), Inches(2.9), W - Inches(1.8), Inches(1.8)),
                    dict(sz=40, colour="FFFFFF", bold=True, algn="l", anchor="t")),
                   ((Inches(0.9), Inches(2.2), W - Inches(1.8), Inches(0.7)),
                    dict(sz=20, colour=mix(P["accent"], "FFFFFF", 0.3), algn="l", anchor="b")),
                   mix(P["dark"], "FFFFFF", 0.5))
        self.shape(s, MSO_SHAPE.OVAL, (-Inches(2.4), H - Inches(3.4), Inches(6.2), Inches(6.2)), P["primary"],
                   alpha=0.45, name="Kit Bg Circle")
        self.shape(s, MSO_SHAPE.OVAL, (Inches(3.3), H - Inches(1.9), Inches(0.8), Inches(0.8)), P["accent"],
                   alpha=0.8, name="Kit Bg Dot")

    def style_geometric(self, L):
        P = self.pal
        self.title_box = (M, Inches(0.35), W - 2 * M - Inches(1.3), Inches(1.0))
        self.content_geometry(L)
        m = self.prs.slide_master
        tri = MSO_SHAPE.RIGHT_TRIANGLE                 # right angle at bottom-left; flipped to the top-right
        self.shape(m, tri, (W - Inches(1.7), 0, Inches(1.7), Inches(1.7)), P["primary"], flip="hv",
                   name="Kit Bg Tri")
        self.shape(m, tri, (W - Inches(0.9), 0, Inches(0.9), Inches(0.9)), P["accent"], flip="hv",
                   name="Kit Bg Tri")
        self.bar(m, M, Inches(1.42))
        t = L["Title Slide"]
        self.cover(t, P["dark"],
                   ((Inches(0.9), Inches(1.4), W * 0.56, Inches(2.3)),
                    dict(sz=42, colour="FFFFFF", bold=True, algn="l", anchor="b")),
                   ((Inches(0.9), Inches(4.0), W * 0.44, Inches(3.0)),
                    dict(sz=20, colour=mix(P["light"], P["dark"], 0.15), algn="l", anchor="t")),
                   mix(P["dark"], "FFFFFF", 0.5))
        self.shape(t, tri, (W * 0.5, H * 0.3, W * 0.5, H * 0.7), P["primary"], flip="h", name="Kit Bg Tri")
        self.shape(t, tri, (W * 0.74, H * 0.62, W * 0.26, H * 0.38), P["accent"], flip="h", name="Kit Bg Tri")
        self.bar(t, Inches(0.9), Inches(3.8), Inches(1.4), Inches(0.07))
        s = L["Section Header"]
        self.cover(s, P["primary"],
                   ((Inches(0.9), Inches(2.9), W * 0.6, Inches(1.8)),
                    dict(sz=40, colour="FFFFFF", bold=True, algn="l", anchor="t")),
                   ((Inches(0.9), Inches(2.2), W * 0.6, Inches(0.7)),
                    dict(sz=20, colour=mix(P["primary"], "FFFFFF", 0.75), algn="l", anchor="b")),
                   mix(P["primary"], "FFFFFF", 0.6))
        self.shape(s, tri, (W * 0.62, H * 0.42, W * 0.38, H * 0.58), P["dark"], flip="h", alpha=0.55,
                   name="Kit Bg Tri")
        self.shape(s, tri, (W * 0.82, H * 0.7, W * 0.18, H * 0.3), P["accent"], flip="h", name="Kit Bg Tri")


def swatches(dest: Path):
    """Picture of every palette (name + its four colours) for choosing one."""
    from PIL import Image, ImageDraw, ImageFont
    import subprocess
    f = subprocess.run(["fc-match", "-f", "%{file}", "Carlito:style=Bold"], capture_output=True, text=True).stdout
    font = ImageFont.truetype(f, 26) if f else ImageFont.load_default()
    cols, cw, ch = 4, 330, 150
    rows = (len(PALETTES) + cols - 1) // cols
    im = Image.new("RGB", (cols * cw + 20, rows * ch + 20), "white")
    d = ImageDraw.Draw(im)
    for k, (name, cs) in enumerate(PALETTES.items()):
        x, y = 20 + (k % cols) * cw, 20 + (k // cols) * ch
        d.text((x, y), name, fill="#222", font=font)
        for j, c in enumerate(cs):
            d.rectangle((x + j * 72, y + 40, x + j * 72 + 66, y + 110), fill="#" + c, outline="#999")
    im.save(dest)
