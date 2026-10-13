"""Catch text copied too closely from the material the report draws on:
the abstracts of cited papers (refs/abstracts.json) and the text files in
sources/ (base paper, notes, chat exports converted to .md/.txt).

Two signals per chapter sentence:
  - a run of RUN or more consecutive words also found in a source (verbatim copy)
  - most of the sentence's 4-word sequences found in one source (close paraphrase)
Similarity checkers like Turnitin flag both; rewrite in your own words and cite.
"""
import json
import re
import subprocess
from pathlib import Path

RUN = 8            # words in a row that count as copied
PARAPHRASE = 0.6   # share of a sentence's 4-grams found in one source


def tokens(text):
    return re.findall(r"[a-z0-9]+", text.lower())


def ngrams(toks, n):
    return {tuple(toks[i:i + n]) for i in range(len(toks) - n + 1)}


def source_texts(rdir):
    """{label: text} for every abstract and every readable file in sources/."""
    out = {}
    ab = rdir / "refs/abstracts.json"
    if ab.exists():
        for key, v in json.loads(ab.read_text()).items():
            if v.get("abstract"):
                out[f"abstract of [@{key}]"] = v["abstract"]
    src = rdir / "sources"
    cache = rdir / "_build/sources-text"
    for f in sorted(src.rglob("*")) if src.exists() else []:
        if not f.is_file():
            continue
        if f.suffix.lower() in (".md", ".txt"):
            out[f"sources/{f.relative_to(src)}"] = f.read_text(errors="ignore")
        elif f.suffix.lower() in (".pdf", ".docx", ".pptx", ".html"):
            cache.mkdir(parents=True, exist_ok=True)
            txt = cache / (f.name + ".txt")
            if not txt.exists() or txt.stat().st_mtime < f.stat().st_mtime:
                kit = Path(__file__).resolve().parent.parent
                r = subprocess.run([str(kit / ".venv/bin/markitdown"), str(f)], capture_output=True, text=True)
                txt.write_text(r.stdout if r.returncode == 0 else "")
            out[f"sources/{f.relative_to(src)}"] = txt.read_text(errors="ignore")
    return out


def sentences(rdir):
    """(file, line, sentence) for chapter prose; code, math and tables skipped."""
    import check
    for f in sorted((rdir / "chapters").glob("*.md")):
        raw = f.read_text()
        cleaned = re.sub(r"```.*?```|\$\$.*?\$\$", lambda m: "\n" * m.group(0).count("\n"), raw, flags=re.S)
        for i, line in enumerate(cleaned.splitlines(), 1):
            if line.lstrip().startswith(("|", "+", "#", "!", ":::", ":")):
                continue
            text = check.prose(line)
            for s in re.split(r"(?<=[.!?])\s+", text):
                if len(tokens(s)) >= RUN:
                    yield f.relative_to(rdir).as_posix(), i, s.strip()


def checks(rdir, final):
    level = "FAIL" if final else "WARN"
    sources = source_texts(rdir)
    if not sources:
        return []
    index = {label: (ngrams(tokens(t), RUN), ngrams(tokens(t), 4)) for label, t in sources.items()}
    out = []
    for f, line, s in sentences(rdir):
        toks = tokens(s)
        runs, quads = ngrams(toks, RUN), ngrams(toks, 4)
        for label, (src_runs, src_quads) in index.items():
            shared = runs & src_runs
            if shared:
                out.append((level, f"{f}:{line} copies {RUN}+ words in a row from {label}: “{' '.join(next(iter(shared)))}…”"))
                break
            if quads and len(quads & src_quads) / len(quads) >= PARAPHRASE:
                out.append((level, f"{f}:{line} closely paraphrases {label}: “{s[:70]}…”"))
                break
    return out
