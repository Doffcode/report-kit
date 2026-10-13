"""Grammar and style check with LanguageTool, run locally (no text leaves the
machine). Writes _build/grammar.md and prints a summary.

    report grammar            British English (closest to Indian academic English)
    report.yml: language: en-US   to switch

Technical writing trips some rules, so a few noisy ones are off, and spelling
"errors" for words in glossary.md, facts.md or refs.bib (jargon, names) are
dropped. Everything else is a suggestion for you to judge, not a build failure.
"""
import json
import re
import socket
import subprocess
import time
import urllib.parse
import urllib.request
from pathlib import Path

KIT = Path(__file__).resolve().parent.parent
LT = next(KIT.glob("tools/LanguageTool-*"), None)
OFF = ["WHITESPACE_RULE", "EN_QUOTES", "DASH_RULE", "COMMA_PARENTHESIS_WHITESPACE",
       "UPPERCASE_SENTENCE_START", "PUNCTUATION_PARAGRAPH_END", "EN_UNPAIRED_BRACKETS",
       "ENGLISH_WORD_REPEAT_BEGINNING_RULE", "SENTENCE_WHITESPACE", "ARROWS", "PLUS_MINUS",
       "OXFORD_SPELLING_Z_NOT_S"]   # -ise is standard British and Indian academic English


def clean_lines(text):
    """Markdown -> plain text with the same line numbers (code/math/tables blanked)."""
    import check
    text = re.sub(r"```.*?```|\$\$.*?\$\$", lambda m: "\n" * m.group(0).count("\n"), text, flags=re.S)
    out = []
    for line in text.splitlines():
        if line.lstrip().startswith(("|", "+-", ":::", ": ", "![")):
            out.append("")
            continue
        line = re.sub(r"^#+\s*", "", line)
        line = re.sub(r"\{[^}]*\}", "", line)
        line = re.sub(r"\[[^\]]*@[^\]]*\]", "", line)        # citations
        line = re.sub(r"@(fig|tbl|eq|alg)-[\w-]+", "Figure 1", line)
        line = re.sub(r"\$[^$]+\$", "x", line)
        line = re.sub(r"\*\*|__|`", "", line)
        line = re.sub(r"^\s*(\d+\.|[-*+])\s+", "", line)
        line = re.sub(r"\[\[TODO[^\]]*\]\]", "", line)
        out.append(line)
    return "\n".join(out)


def known_words(rdir):
    words = set()
    for name in ("glossary.md", "facts.md", "refs.bib"):
        f = rdir / name
        if f.exists():
            words |= set(re.findall(r"[A-Za-z][A-Za-z0-9\-]+", f.read_text()))
    return {w.lower() for w in words}


def free_port():
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


def run(rdir, cfg):
    if LT is None:
        raise SystemExit("LanguageTool is not installed in ~/report-kit/tools")
    lang = cfg.get("language", "en-GB")
    port = free_port()
    import shutil
    java = shutil.which("java", path=f"{KIT / 'bin'}:{__import__('os').environ.get('PATH', '')}")
    if not java:
        raise SystemExit("Java not found: re-run ~/report-kit/install.sh (it installs a private Java)")
    server = subprocess.Popen([java, "-cp", str(LT / "languagetool-server.jar"),
                               "org.languagetool.server.HTTPServer", "--port", str(port)],
                              stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        url = f"http://127.0.0.1:{port}/v2/check"
        for _ in range(120):
            try:
                urllib.request.urlopen(f"http://127.0.0.1:{port}/v2/languages", timeout=1)
                break
            except Exception:
                time.sleep(0.25)
        known = known_words(rdir)
        files = sorted((rdir / "front").glob("*.md")) + sorted((rdir / "chapters").glob("*.md"))
        report, total = ["# Grammar and style", "",
                         f"LanguageTool {LT.name.split('-')[-1]}, {lang}. Suggestions to judge, not errors to obey.", ""], 0
        for f in files:
            text = clean_lines(f.read_text())
            data = urllib.parse.urlencode({"language": lang, "text": text,
                                           "disabledRules": ",".join(OFF)}).encode()
            res = json.loads(urllib.request.urlopen(url, data=data, timeout=120).read())
            starts = [0]
            for line in text.splitlines(keepends=True):
                starts.append(starts[-1] + len(line))
            items = []
            for m in res.get("matches", []):
                word = text[m["offset"]:m["offset"] + m["length"]]
                if m["rule"]["issueType"] == "misspelling" and (
                        word.lower() in known or re.search(r"[A-Z].*[A-Z]|\d", word)):
                    continue
                line = next(i for i in range(len(starts) - 1) if starts[i + 1] > m["offset"]) + 1
                fix = ", ".join(r["value"] for r in m.get("replacements", [])[:3])
                items.append(f"- line {line}: “{word}” — {m['message']}" + (f" → {fix}" if fix else ""))
            if items:
                report += [f"## {f.relative_to(rdir).as_posix()} ({len(items)})", ""] + items + [""]
                total += len(items)
                print(f"  {f.relative_to(rdir).as_posix()}: {len(items)}")
        out = rdir / "_build/grammar.md"
        out.parent.mkdir(exist_ok=True)
        out.write_text("\n".join(report))
        print(f"\n  {total} suggestion(s) → {out.relative_to(rdir.parent)}")
    finally:
        server.terminate()
        server.wait(timeout=20)
