"""Report formats ("profiles").

A profile is a YAML file in ~/report-kit/profiles/ (or a report's own
profile.yml) describing everything format-specific: paper, margins, font,
spacing, chapter heading style, cover lines, certificate/declaration
wording and signatures. `extends:` inherits from another profile; a
report's report.yml can override any key under `format:`.

The build fills the {placeholders} once and hands the same result to the
Typst template (as kit-profile.typ) and to the DOCX post-processor, so both
outputs always agree.
"""
import re
from pathlib import Path

import yaml

KIT = Path(__file__).resolve().parent.parent
PROFILES = KIT / "profiles"
PLACEHOLDER = re.compile(r"\{([a-z_]+)\}")


def deep_merge(base, over):
    out = dict(base)
    for k, v in (over or {}).items():
        out[k] = deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def find_profile(name, rdir):
    for p in (rdir / f"{name}.yml", rdir / "profiles" / f"{name}.yml", PROFILES / f"{name}.yml"):
        if p.exists():
            return p
    have = sorted(p.stem for p in PROFILES.glob("*.yml"))
    raise SystemExit(f"profile {name!r} not found (kit has: {', '.join(have)})")


def load(name, rdir, overrides=None, _seen=()):
    if name in _seen:
        raise SystemExit(f"profile {name!r} extends itself")
    data = yaml.safe_load(find_profile(name, rdir).read_text()) or {}
    parent = data.pop("extends", None)
    base = load(parent, rdir, None, _seen + (name,)) if parent else {}
    return deep_merge(deep_merge(base, data), overrides or {})


def values(cfg, prof):
    """Every {placeholder} a profile may use."""
    me = cfg.get("me") or {}
    g, h = me.get("guide") or {}, me.get("hod") or {}
    intern = cfg.get("internship") or {}
    date = str(cfg.get("date") or "")
    v = {
        "title": cfg.get("title"), "report_type": cfg.get("subtitle") or prof.get("report_type"),
        "doc_kind": prof.get("doc_kind"), "date": date,
        "year": (re.search(r"\d{4}", date) or [None])[0] if date else None,
        "name": me.get("name"), "roll": me.get("roll"), "degree": me.get("degree"),
        "branch": me.get("branch"), "department": me.get("department"), "college": me.get("college"),
        "place": me.get("place"), "guide": g.get("name"), "guide_title": g.get("title"),
        "hod": h.get("name"), "hod_title": h.get("title"),
        "company": intern.get("company"), "company_place": intern.get("place"),
        "period": intern.get("period"), "mentor": intern.get("mentor"),
        "mentor_title": intern.get("mentor_title"),
    }
    return {k: str(x) for k, x in v.items() if x not in (None, "")}


# ── filling text for each output ─────────────────────────────
TYPST_SPECIAL = re.compile(r"([\\#*_`$<>@\[\]~=/-])")


def typst_escape(s):
    return TYPST_SPECIAL.sub(r"\\\1", s)


def fill(text, vals, target, missing):
    """Replace {placeholders}. target 'typst' escapes values for Typst markup and
    marks gaps with a highlight; target 'docx' leaves values raw and marks gaps
    as [[TODO: ...]] (docx_fix highlights those)."""
    def sub(m):
        key = m.group(1)
        if key in vals:
            return typst_escape(vals[key]) if target == "typst" else vals[key]
        missing.add(key)
        return f"#highlight[TODO: {key}]" if target == "typst" else f"[[TODO: {key}]]"
    return PLACEHOLDER.sub(sub, str(text))


def resolve(prof, cfg, rdir, target):
    """Profile with all text filled for one output. Returns (profile, missing keys)."""
    vals = values(cfg, prof)
    missing = set()
    out = dict(prof)
    out["cover"] = []
    for item in prof.get("cover") or []:
        item = dict(item)
        if "text" in item:
            # Drop a line whose placeholders are all empty (e.g. no branch): no "in " orphan.
            keys = PLACEHOLDER.findall(item["text"])
            if keys and not any(k in vals for k in keys) and item.get("optional", True):
                continue
            item["text"] = fill(item["text"], vals, target, missing)
        if "logo" in item:
            logo = (cfg.get("me") or {}).get("logo")
            if not logo:
                continue
            item["path"] = logo
        out["cover"].append(item)
    out["pages"] = []
    for page in prof.get("pages") or []:
        page = dict(page)
        if "text" in page:
            page["text"] = fill(" ".join(str(page["text"]).split()), vals, target, missing)
        if "image" in page:
            found = next((f"{page['image']}{ext}" for ext in (".png", ".jpg", ".jpeg")
                          if (rdir / f"{page['image']}{ext}").exists()), None)
            if found:
                page["image"] = found
            else:
                missing.add(f"{page.pop('image')}.png")
                desc = fill(page.get("missing", "image"), vals, target, missing)
                page["text"] = f"#highlight[TODO: {desc}]" if target == "typst" else f"[[TODO: {desc}]]"
        sigs = []
        for who in page.get("sign") or []:
            s = (prof.get("signatories") or {}).get(who, {"name": f"{{{who}}}", "role": ""})
            sigs.append({"name": fill(s["name"], vals, target, missing),
                         "role": fill(s["role"], vals, target, missing)})
        page["sign"] = sigs
        page["place"] = fill("{place}", vals, target, missing) if page.get("place_date") else None
        out["pages"].append(page)
    return out, missing


def length_pt(v):
    """'0.8cm' / '1.5in' / '12pt' / '3mm' -> points."""
    m = re.fullmatch(r"(-?[\d.]+)(pt|cm|mm|in)", str(v).strip())
    if not m:
        return 0.0
    return float(m.group(1)) * {"pt": 1, "cm": 28.3465, "mm": 2.83465, "in": 72}[m.group(2)]


def margins_pt(prof):
    m = (prof.get("page") or {}).get("margin") or {}
    return {k: length_pt(m.get(k, d)) for k, d in
            (("left", "1.5in"), ("right", "1in"), ("top", "1in"), ("bottom", "1in"))}


# ── Typst serialisation ──────────────────────────────────────
LENGTH = re.compile(r"^-?\d+(\.\d+)?(pt|cm|mm|in|em|fr|%)$")


def to_typst(v):
    if isinstance(v, bool):
        return "true" if v else "false"
    if v is None:
        return "none"
    if isinstance(v, (int, float)):
        return repr(v)
    if isinstance(v, str):
        if LENGTH.match(v):
            return v
        return '"' + v.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n") + '"'
    if isinstance(v, list):
        return "(" + "".join(to_typst(x) + ", " for x in v) + ")"
    if isinstance(v, dict):
        if not v:
            return "(:)"
        return "(" + ", ".join(f'"{k}": {to_typst(x)}' for k, x in v.items()) + ")"
    raise TypeError(type(v))


def leading_em(spacing):
    """Typst leading giving Word's 'multiple' line spacing for Times-like fonts,
    with text edges at ascender/descender (line box ≈ 1.107em, Word single ≈ 1.149em)."""
    return round(spacing * 1.149 - 1.107, 3)


# Metric-compatible stand-ins: same character widths, so page breaks don't move
# if the named font is missing on some machine.
FALLBACK = {"times new roman": "Liberation Serif", "arial": "Liberation Sans",
            "helvetica": "Liberation Sans", "courier new": "Liberation Mono"}


def write_typst(prof, rdir):
    p = dict(prof)
    fam = (prof.get("font") or {}).get("family", "Times New Roman")
    p["font"] = {**prof.get("font", {}), "fallback": FALLBACK.get(fam.lower(), "Liberation Serif")}
    p["leading"] = f"{leading_em(float(prof.get('spacing', 1.5)))}em"
    (rdir / "kit-profile.typ").write_text(
        "// Generated by `report build` from the format profile. Do not edit.\n"
        f"#let profile = {to_typst(p)}\n")
