"""Checks run on every build. Each returns (level, message) tuples.

FAIL blocks a --final build; in a draft build, content gaps (TODOs, numbers
missing from facts.md, unreferenced figures) are only warnings so you can
still read the PDF.
"""
import re
from pathlib import Path

import pymupdf

ALLOWED_FONTS = ("TimesNewRoman", "Arial", "NewCMMath", "NewComputerModernMath", "DejaVuSansMono",
                 "LiberationSerif", "LiberationSans")
NUM_RE = re.compile(r"(?<![\w.@-])\d{1,3}(?:,\d{3})+(?:\.\d+)?%?|(?<![\w.@-])\d+(?:\.\d+)?%?")
ACRONYM_RE = re.compile(r"\b([A-Z][A-Z0-9]{1,}[a-z]?)\b")
NOT_ACRONYMS = {"II", "III", "IV", "VI", "VII", "VIII", "IX", "XI", "XII", "TODO", "OK", "AND", "OR", "NOT"}


def read_glossary(rdir: Path) -> dict:
    g = rdir / "glossary.md"
    if not g.exists():
        return {}
    rows = {}
    for line in g.read_text().splitlines():
        m = re.match(r"^\s*[-*]\s*\**`?([A-Za-z][A-Za-z0-9/&-]*)`?\**\s*[:—–]\s*(.+?)\s*$", line)
        if m:
            rows[m.group(1)] = m.group(2)
    return rows


def content_files(rdir):
    return sorted((rdir / "front").glob("*.md")) + sorted((rdir / "chapters").glob("*.md"))


def prose(text: str) -> str:
    """Strip code, math, attributes, citations, refs, image targets: what's left is prose."""
    text = re.sub(r"```.*?```", " ", text, flags=re.S)
    text = re.sub(r"\$\$.*?\$\$", " ", text, flags=re.S)
    text = re.sub(r"\$[^$\n]+\$", " ", text)
    text = re.sub(r"`[^`]*`", " ", text)
    text = re.sub(r"\{[^}]*\}", " ", text)                      # {#fig-x width=80%}
    text = re.sub(r"\]\([^)]*\)", "]", text)                     # link/image targets
    text = re.sub(r"\[[^\]]*@[^\]]*\]", " ", text)               # [@cite; @cite2]
    text = re.sub(r"@[\w:-]+", " ", text)                        # @fig-x
    text = re.sub(r"^\s*(\d+\.|[-*+]|#+)\s", " ", text, flags=re.M)  # list markers, headings
    text = re.sub(r"^\s*[|+][-=:+|\s]*$", " ", text, flags=re.M)     # table rules
    text = re.sub(r"\[\[TODO[^\]]*\]\]", " ", text)
    return text


def norm_num(s: str) -> str:
    return s.replace(",", "").rstrip("%")


def run_all(rdir: Path, cfg: dict, final: bool, logs: str = ""):
    out = []
    gap = "FAIL" if final else "WARN"
    files = content_files(rdir)
    texts = {f: f.read_text() for f in files}
    rel = lambda f: f.relative_to(rdir).as_posix()  # noqa: E731

    # 1. Build-log problems: unresolved cross-references and citations.
    for line in logs.splitlines():
        if line.startswith("REPORTKIT-FAIL "):
            out.append(("FAIL", line[len("REPORTKIT-FAIL "):]))
        elif re.search(r"Unable to resolve crossref|citation .* not found|Could not convert image", line):
            out.append(("FAIL", line.strip()[:160]))
        elif "WARNING" in line and "rsvg" not in line:
            out.append(("WARN", line.strip()[:160]))

    # 2. Leftover TODOs.
    for f, t in texts.items():
        for i, line in enumerate(t.splitlines(), 1):
            for m in re.finditer(r"\[\[TODO:?\s*([^\]]*)\]\]", line):
                out.append((gap, f"TODO in {rel(f)}:{i}: {m.group(1)[:80]}"))

    # 3. Figures/tables/algorithms: defined, referenced, files present.
    all_text = "\n".join(texts.values())
    defined = re.findall(r"#((?:fig|tbl|alg)-[\w-]+)", all_text)
    referenced = set(re.findall(r"@((?:fig|tbl|alg|eq)-[\w-]+)", all_text))
    for label in defined:
        if label not in referenced:
            out.append((gap, f"{label} is never referenced in the text (add “see @{label}”)"))
    for f, t in texts.items():
        for m in re.finditer(r"!\[[^\]]*\]\((figures/[^)\s]+)\)", t):
            stem = rdir / re.sub(r"\.(svg|png|pdf|jpg)$", "", m.group(1))
            if not (stem.with_suffix(".svg").exists() or stem.with_suffix(".png").exists()):
                out.append(("FAIL", f"missing image {m.group(1)} (used in {rel(f)})"))

    # 4. Numbers in prose must be traceable to facts.md.
    facts = (rdir / "facts.md").read_text() if (rdir / "facts.md").exists() else ""
    known = {norm_num(n) for n in NUM_RE.findall(facts)}
    for f in (rdir / "chapters").glob("*.md"):
        for n in sorted(set(NUM_RE.findall(prose(texts[f])))):
            v = norm_num(n)
            if v in known or (v.isdigit() and int(v) <= 10):
                continue
            out.append((gap, f"number {n} in {rel(f)} is not in facts.md"))

    # 5. Abbreviations: defined in glossary, expanded at first use.
    gloss = read_glossary(rdir)
    ignore = NOT_ACRONYMS | set(cfg.get("not_abbreviations") or [])
    chapters_md = "\n".join(texts[f] for f in sorted((rdir / "chapters").glob("*.md")))
    chapters_md = re.sub(r"^#+ .*$", " ", chapters_md, flags=re.M)   # headings don't count as first use
    body = prose(chapters_md)
    seen = []
    for m in ACRONYM_RE.finditer(body):
        a = m.group(1).rstrip("s")
        if a not in seen and a not in ignore and not re.fullmatch(r"[LR]\d", a):
            seen.append(a)
            if a not in gloss:
                out.append(("WARN", f"abbreviation {a} is not in glossary.md"))
            elif f"({a})" not in body[: m.end() + 1]:
                out.append(("WARN", f"{a} is not expanded at first use: write “{gloss[a]} ({a})”"))

    # 6. PDF: unresolved refs, fonts, page count.
    pdf = rdir / "out/report.pdf"
    if pdf.exists():
        doc = pymupdf.open(pdf)
        text = "".join(p.get_text() for p in doc)
        if "?@" in text or "??" in text:
            out.append(("FAIL", "unresolved reference (?? or ?@) in the PDF"))
        fonts = set()
        for p in doc:
            for f in p.get_fonts():
                if f[2] == "Type3":
                    out.append(("FAIL", f"Type 3 font {f[3]} on page {p.number + 1} (IEEE rejects these)"))
                fonts.add(f[3].split("+")[-1])
        for f in sorted(fonts):
            if not f.replace("-", "").replace(" ", "").startswith(ALLOWED_FONTS):
                out.append(("WARN", f"unexpected font in PDF: {f} (fallback?)"))
        target = cfg.get("target_pages")
        if target:
            lo, hi = target if isinstance(target, list) else (0, target)
            if not lo <= len(doc) <= hi:
                out.append((gap, f"PDF has {len(doc)} pages; target is {lo}–{hi}"))
        out.append(("INFO", f"PDF: {len(doc)} pages, fonts: {', '.join(sorted(fonts))}"))
    return out
