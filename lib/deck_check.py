"""Content checks for slides.md (layout problems are found while building)."""
import re
from pathlib import Path

import check

LABEL_TITLES = {"introduction", "results", "methodology", "method", "methods",
                "discussion", "evaluation", "experiments", "experimental results", "implementation",
                "proposed system", "proposed method", "system design", "architecture",
                "literature survey", "literature review", "related work", "analysis", "overview", "outline"}


def run(sdir: Path, cfg: dict, slides: list, final: bool):
    out = []
    gap = "FAIL" if final else "WARN"
    rdir = cfg.get("_report")
    md = (sdir / "slides.md").read_text()

    for i, line in enumerate(md.splitlines(), 1):
        for m in re.finditer(r"\[\[TODO:?\s*([^\]]*)\]\]", line):
            out.append((gap, f"TODO in slides.md:{i}: {m.group(1)[:80] or line.strip()[:80]}"))

    content = [s for s in slides if s["level"] == 2 and not set(s["classes"]) & {"references", "qa", "backup"}]
    for s in content:
        t = re.sub(r"\[\[TODO[^\]]*\]\]", "", s["title"]).strip().lower()
        if t in LABEL_TITLES:
            out.append(("WARN", f"slides.md:{s['line']}: title “{s['title']}” only names the section; "
                                f"state the point (e.g. “FedProx beats FedAvg on non-IID data”)"))
        if not s["notes"]:
            out.append((gap if final else None, f"slides.md:{s['line']}: “{s['title'][:40]}” has no ::: notes"))
    out = [o for o in out if o[0]]
    missing_notes = sum(1 for s in content if not s["notes"])
    if missing_notes and not final:
        out.append(("INFO", f"{missing_notes} slide(s) without speaker notes (required with --final)"))

    minutes = cfg.get("talk_minutes")
    if minutes:
        n = len(content)
        lo, hi = round(minutes * 0.6), round(minutes * 1.3)
        if not lo <= n <= hi:
            out.append(("WARN", f"{n} content slides for a {minutes}-minute talk (about {lo}–{hi} is comfortable)"))

    # numbers must come from facts.md, abbreviations from glossary.md (slides/ first, then report/)
    body = "\n".join(re.sub(r"^#+ .*$", " ", ln) for ln in md.splitlines() if not ln.startswith("<!--"))
    body = re.sub(r"<!--.*?-->", " ", body, flags=re.S)
    body = re.sub(r"^:::\s*notes.*?^:::\s*$", " ", body, flags=re.S | re.M)
    text = check.prose(body)
    facts = "".join((d / "facts.md").read_text() for d in (sdir, rdir) if d and (d / "facts.md").exists())
    if facts:
        known = {check.norm_num(n) for n in check.NUM_RE.findall(facts)}
        for n in sorted(set(check.NUM_RE.findall(text))):
            v = check.norm_num(n)
            if v not in known and not (v.isdigit() and int(v) <= 10):
                out.append((gap, f"number {n} in slides.md is not in facts.md"))
    gloss = {}
    for d in (rdir, sdir):
        if d:
            gloss.update(check.read_glossary(d))
    if gloss:
        seen = []
        ignore = check.NOT_ACRONYMS | set(cfg.get("not_abbreviations") or [])
        for m in check.ACRONYM_RE.finditer(text):
            a = m.group(1).rstrip("s")
            if a in seen or a in ignore:
                continue
            seen.append(a)
            if a in gloss and f"({a})" not in text[: m.end() + 1] and gloss[a].lower() not in text.lower():
                out.append(("WARN", f"{a} is not spelled out on the slides: write “{gloss[a]} ({a})” once"))
    return out
