"""report: turn a report/ folder into report.pdf + report.docx.

    report init [DIR] [--profile thesis]   create a report/ skeleton (default ./report)
    report build [--final] [--pdf-only]
    report check [--final]     checks only, on the last build
    report inspect             layout inspection only, on the last build
    report cite add "<title|DOI|arXiv>"   fetch a reference into refs.bib
    report cite verify [--fix]           check refs.bib against Crossref/DBLP/OpenAlex/arXiv
    report cite review                   citing sentences next to abstracts
    report style-extract sample.pdf --name <college>   draft a format profile from a sample PDF
    report grammar                       local grammar/style suggestions → _build/grammar.md

Run from inside a report/ folder (or its parent project folder).
"""
import argparse
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

import yaml

KIT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(KIT / "lib"))
BIN = KIT / "bin"
ENV = {**os.environ, "PATH": f"{BIN}:{os.environ['PATH']}"}
TEMPLATE = KIT / "templates/report"

import check  # noqa: E402
import similarity  # noqa: E402
import profile as fmt  # noqa: E402
import cite  # noqa: E402
import inspect_layout  # noqa: E402


def say(msg):
    print(f"\033[1m▸\033[0m {msg}", flush=True)


def find_report_dir(start: Path) -> Path:
    for d in (start, start / "report"):
        if (d / "report.yml").exists():
            return d
    sys.exit("No report.yml here or in ./report. Run `report init` first.")


def load_config(rdir: Path) -> dict:
    cfg = yaml.safe_load((rdir / "report.yml").read_text()) or {}
    me_path = rdir / "me.yml" if (rdir / "me.yml").exists() else KIT / "me.yml"
    me = (yaml.safe_load(me_path.read_text()) or {}) if me_path.exists() else {}
    cfg["me"] = {**me.get("me", {}), **(cfg.get("me") or {})}
    cfg.setdefault("date", time.strftime("%B %Y"))
    logo = cfg["me"].get("logo")
    if logo and not (rdir / logo).exists():
        shared = me_path.parent / logo
        if shared.exists():   # Typst only reads files inside the report folder
            local = rdir / "assets" / f"kit-{shared.name}"
            local.parent.mkdir(exist_ok=True)
            shutil.copy2(shared, local)
            cfg["me"]["logo"] = local.relative_to(rdir).as_posix()
        else:
            print(f"warning: logo {logo} not found; cover will have no logo")
            cfg["me"].pop("logo")
    return cfg


# ── diagrams ─────────────────────────────────────────────────
def stale(src: Path, outs):
    return any(not o.exists() or o.stat().st_mtime < src.stat().st_mtime for o in outs)


def render_diagrams(rdir: Path):
    figs = rdir / "figures"
    figs.mkdir(exist_ok=True)
    mm_cfg = KIT / "styles/mermaid.json"
    tt = Path.home() / ".local/share/fonts/msttcore"
    d2_fonts = ["--font-regular", str(tt / "times.ttf"), "--font-bold", str(tt / "timesbd.ttf"),
                "--font-italic", str(tt / "timesi.ttf")]
    n = 0
    for src in sorted((rdir / "diagrams").glob("*")):
        svg, png = figs / f"{src.stem}.svg", figs / f"{src.stem}.png"
        if not stale(src, [svg, png]):
            continue
        if src.suffix == ".mmd":
            for out, extra in ((svg, []), (png, ["-s", "3"])):
                run(["mmdc", "-i", src, "-o", out, "-b", "white", "-c", mm_cfg, *extra], quiet=True)
        elif src.suffix == ".d2":
            run(["d2", "--layout=elk", "--pad", "20", *d2_fonts, src, svg], quiet=True)
            run(["d2", "--layout=elk", "--pad", "20", "--scale", "3", *d2_fonts, src, png], quiet=True)
        else:
            continue
        n += 1
    for src in sorted((rdir / "charts").glob("*.py")):
        marker = figs / f".{src.stem}.done"
        if stale(src, [marker]):
            run([KIT / ".venv/bin/python", src], cwd=rdir, quiet=True)
            marker.touch()
            n += 1
    if n:
        say(f"rendered {n} diagram/chart source(s)")


def run(cmd, cwd=None, quiet=False):
    r = subprocess.run([str(c) for c in cmd], cwd=cwd, env=ENV, capture_output=True, text=True)
    if r.returncode != 0:
        sys.stderr.write(r.stdout + r.stderr)
        sys.exit(f"failed: {' '.join(map(str, cmd))}")
    return r


# ── document assembly ────────────────────────────────────────
def md_files(rdir, sub):
    return sorted(p.relative_to(rdir).as_posix() for p in (rdir / sub).glob("*.md")) if (rdir / sub).exists() else []


def abbreviations_md(rdir: Path):
    rows = check.read_glossary(rdir)
    if not rows:
        return ""
    lines = ["# List of Abbreviations {.unnumbered}", "", "| Abbreviation | Expansion |", "|:--|:--|"]
    lines += [f"| {a} | {e} |" for a, e in sorted(rows.items())]
    return "\n".join(lines) + "\n"


def write_qmd(rdir: Path, cfg: dict) -> Path:
    def kit(p):  # Quarto resolves partials/filters relative to the document
        return os.path.relpath(KIT / p, rdir)

    style = cfg.get("citation_style", "ieee")
    if not (KIT / f"styles/{style}.csl").exists():
        sys.exit(f"citation_style {style!r}: no styles/{style}.csl in the kit (have: ieee, apa)")
    (rdir / "refs").mkdir(exist_ok=True)
    shutil.copy2(KIT / f"styles/{style}.csl", rdir / f"refs/{style}.csl")   # Typst reads only inside the report

    prof = load_profile(rdir, cfg)
    typst_prof, _ = fmt.resolve(prof, cfg, rdir, "typst")
    typst_prof["doc_title"] = cfg.get("title", "")
    fmt.write_typst(typst_prof, rdir)
    caps = prof.get("captions") or {}
    meta = {
        "title": cfg.get("title", "Untitled Report"),
        "filters": [{"path": kit("filters/algorithm.lua"), "at": "pre-ast"},
                    {"path": kit("filters/math.lua"), "at": "pre-ast"},
                    {"path": kit("filters/tables.lua"), "at": "pre-ast"},
                    {"path": kit("filters/todo.lua"), "at": "pre-ast"}],
        "crossref": {"fig-prefix": caps.get("figure", "Figure"), "fig-title": caps.get("figure", "Figure"),
                     "tbl-prefix": caps.get("table", "Table"), "tbl-title": caps.get("table", "Table"),
                     "custom": [{"kind": "float", "key": "alg", "reference-prefix": "Algorithm",
                                 "caption-prefix": "Algorithm", "caption-location": "top"}]},
        "format": {
            "typst": {
                "keep-typ": True,
                "csl": f"refs/{style}.csl",   # same CSL file as the DOCX, copied inside the report root (no "_": pandoc escapes it)
                "bibliography": cite.write_typst_bib(rdir, style),
                "default-image-extension": "svg",
                "template-partials": [kit("templates/report/typst-template.typ"), kit("templates/report/typst-show.typ")],
            },
            "docx": {
                "default-image-extension": "png",
                "reference-doc": kit("templates/report/reference.docx"),
                "csl": kit(f"styles/{style}.csl"),
                "bibliography": "refs.bib",
                "number-sections": True,
                "crossref": {"chapters": True},
            },
        },
    }
    # Bibliography is set per format only: Quarto merges a top-level one with format ones.
    has_refs = (rdir / "refs.bib").exists() and (rdir / "refs.bib").read_text().strip()
    if not has_refs:
        for f in ("typst", "docx"):
            meta["format"][f].pop("bibliography", None)
    parts = ["---", yaml.safe_dump(meta, sort_keys=False, allow_unicode=True).strip(), "---", ""]
    parts += [f"{{{{< include {f} >}}}}\n" for f in cfg.get("front") or md_files(rdir, "front")]
    parts += ["```{=typst}", "#frontlists()", "```", ""]
    abbr = abbreviations_md(rdir) if prof.get("abbreviations", True) else ""
    if abbr:
        parts += [abbr]
    parts += ["```{=typst}", "#show: mainmatter", "```", ""]
    chapters = cfg.get("chapters") or md_files(rdir, "chapters")
    if not chapters:
        sys.exit("No chapters: add markdown files to report/chapters/.")
    parts += [f"{{{{< include {f} >}}}}\n" for f in chapters]
    if has_refs:
        parts += ['::: {.content-visible when-format="docx"}', "# References {.unnumbered}", "",
                  "::: {#refs}", ":::", ":::", ""]
    qmd = rdir / "_report.qmd"
    qmd.write_text("\n".join(parts))
    return qmd


def load_profile(rdir, cfg):
    """The report's format: profile named in report.yml (default: project),
    with any `format:` overrides from report.yml applied on top."""
    return fmt.load(cfg.get("profile", "project"), rdir, cfg.get("format"))


def render(rdir: Path, qmd: Path, fmt: str):
    r = subprocess.run(["quarto", "render", qmd.name, "--to", fmt], cwd=rdir, env=ENV,
                       capture_output=True, text=True)
    log = r.stdout + r.stderr
    (rdir / "_build").mkdir(exist_ok=True)
    (rdir / f"_build/quarto-{fmt}.log").write_text(log)
    if r.returncode != 0:
        sys.stderr.write(log[-4000:])
        sys.exit(f"{fmt} build failed (full log: report/_build/quarto-{fmt}.log)")
    return log


def build(rdir: Path, final: bool, pdf_only: bool):
    t0 = time.time()
    cfg = load_config(rdir)
    render_diagrams(rdir)
    qmd = write_qmd(rdir, cfg)
    out = rdir / "out"
    out.mkdir(exist_ok=True)

    say("building PDF")
    logs = render(rdir, qmd, "typst")
    shutil.move(rdir / "_report.pdf", out / "report.pdf")
    if not pdf_only:
        say("building DOCX")
        logs += render(rdir, qmd, "docx")
        shutil.move(rdir / "_report.docx", out / "report.docx")
        import docx_fix
        docx_prof, _ = fmt.resolve(load_profile(rdir, cfg), cfg, rdir, "docx")
        docx_fix.fix(out / "report.docx", docx_prof, rdir)
        (rdir / "_build/docx-preview.pdf").unlink(missing_ok=True)
        has_uno = shutil.which("soffice") and subprocess.run(
            ["python3", "-c", "import uno"], capture_output=True).returncode == 0
        if not has_uno:
            # No LibreOffice here: ask Word to refresh TOC/LOF/LOT when the file is opened.
            docx_fix.update_fields_on_open(out / "report.docx")
            logs += ("\nWARNING LibreOffice (with python3-uno) not found: TOC/LOF/LOT in the DOCX will fill when "
                     "Word opens it (click Yes), and the Word layout check is skipped")
            r = None
        else:
            r = subprocess.run(["python3", KIT / "lib/lo_refresh.py", out / "report.docx",
                                rdir / "_build/docx-preview.pdf"], capture_output=True, text=True, timeout=180)
        if r is not None and r.returncode != 0:
            (rdir / "_build/lo-refresh.log").write_text(r.stdout + r.stderr)
            lo_error = "DOCX TOC/LOF/LOT not filled: LibreOffice refresh failed (see report/_build/lo-refresh.log)"
            logs += "\nREPORTKIT-FAIL " + lo_error
    say(f"built in {time.time() - t0:.1f}s → {out.relative_to(rdir.parent) if rdir.parent else out}")
    ok = report_quality(rdir, cfg, final, logs)
    if not ok:
        sys.exit(1)


def report_quality(rdir, cfg, final, logs=""):
    problems = (check.run_all(rdir, cfg, final, logs) + cite.checks(rdir, final)
                + similarity.checks(rdir, final) + front_checks(rdir, cfg, final))
    layout = inspect_layout.inspect(rdir, cfg)
    print()
    for level, msg in problems + layout:
        mark = {"FAIL": "\033[31m✗\033[0m", "WARN": "\033[33m!\033[0m", "INFO": "·"}[level]
        print(f"  {mark} {msg}")
    fails = [p for p in problems + layout if p[0] == "FAIL"]
    warns = [p for p in problems + layout if p[0] == "WARN"]
    print(f"\n  {len(fails)} failure(s), {len(warns)} warning(s)"
          + ("" if final else "  [draft build: use --final before submitting]"))
    return not fails


def front_checks(rdir, cfg, final):
    """Details the cover/certificate need but don't have, and one-shot assumptions."""
    out = []
    _, missing = fmt.resolve(load_profile(rdir, cfg), cfg, rdir, "docx")
    if missing:
        out.append(("FAIL" if final else "WARN",
                    f"cover/front pages need: {', '.join(sorted(missing))} "
                    "(set in ~/report-kit/me.yml, report.yml, or add the file)"))
    a = rdir / "assumptions.md"
    if a.exists():
        n = sum(1 for line in a.read_text().splitlines() if line.lstrip().startswith(("-", "*")))
        if n:
            out.append(("WARN" if not final else "FAIL",
                        f"{n} assumption(s) made without asking you: review report/assumptions.md, "
                        "then delete the lines you've confirmed"))
    return out


# ── init ─────────────────────────────────────────────────────
SKELETON_YML = """\
# Report settings. Your personal details live in ~/report-kit/me.yml.
title: "{title}"
profile: project            # project | thesis | internship | seminar (see ~/report-kit/profiles)
# date: "May 2027"          # defaults to the current month
# citation_style: ieee      # ieee | apa
# format:                   # override any profile setting, e.g.
#   spacing: 2.0
#   chapter: {{style: inline}}
# internship: {{company: ..., place: ..., period: "May–July 2026", mentor: ..., mentor_title: ...}}
# target_pages: [40, 60]    # checked on the PDF
# chapters: [01-intro.md]   # optional explicit order; default = chapters/*.md sorted
"""


def init(target: Path, profile_name="project"):
    if (target / "report.yml").exists():
        sys.exit(f"{target}/report.yml already exists.")
    fmt.find_profile(profile_name, target)   # fail early on a typo, before creating anything
    for sub in ("front", "chapters", "diagrams", "charts", "figures", "assets", "sources"):
        (target / sub).mkdir(parents=True, exist_ok=True)
    title = target.resolve().parent.name.replace("-", " ").replace("_", " ").title()
    (target / "report.yml").write_text(SKELETON_YML.format(title=title).replace(
        "profile: project ", f"profile: {profile_name} ".ljust(len("profile: project "))))
    defaults = {
        "front/01-acknowledgement.md": "# Acknowledgement {.unnumbered}\n\n[[TODO: acknowledgement]]\n",
        "front/02-abstract.md": "# Abstract {.unnumbered}\n\n[[TODO: abstract]]\n\n**Keywords:** [[TODO: keywords]]\n",
        "facts.md": "# Facts\n\nEvery number used in the report, with its source.\n\n",
        "glossary.md": "# Glossary\n\nOne abbreviation per line: `- ABBR: Expansion`\n\n",
        "refs.bib": "",
        ".gitignore": "_report.*\nkit-profile.typ\n_build/\n*_files/\n.quarto/\nfigures/.*.done\nrefs/typst.bib\nrefs/*.csl\n",
    }
    for name, text in defaults.items():
        if not (target / name).exists():
            (target / name).write_text(text)
    print(f"created {target}/ — put source material in {target}/sources/")


def main():
    ap = argparse.ArgumentParser(prog="report", description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("init")
    p.add_argument("dir", nargs="?", default="report")
    p.add_argument("--profile", default="project", help="project | thesis | internship | seminar | <your own>")
    p = sub.add_parser("build")
    p.add_argument("--final", action="store_true")
    p.add_argument("--pdf-only", action="store_true")
    p = sub.add_parser("check")
    p.add_argument("--final", action="store_true")
    sub.add_parser("inspect")
    sub.add_parser("grammar", help="local LanguageTool grammar/style pass → _build/grammar.md")
    p = sub.add_parser("style-extract", help="draft a format profile from a sample report PDF")
    p.add_argument("pdf")
    p.add_argument("--name", required=True, help="profile name, e.g. nitt-thesis")
    p.add_argument("--force", action="store_true")
    p = sub.add_parser("cite", help="add / verify / review references")
    p.add_argument("action", choices=["add", "verify", "review"])
    p.add_argument("queries", nargs="*", help="for add: titles, DOIs or arXiv ids")
    p.add_argument("--fix", action="store_true", help="for verify: replace wrong fields with the real record")
    a = ap.parse_args()

    if a.cmd == "init":
        return init(Path(a.dir), a.profile)
    if a.cmd == "style-extract":
        import style_extract
        return style_extract.main(a.pdf, a.name, a.force)
    rdir = find_report_dir(Path.cwd())
    if a.cmd == "build":
        build(rdir, a.final, a.pdf_only)
    elif a.cmd == "check":
        sys.exit(0 if report_quality(rdir, load_config(rdir), a.final) else 1)
    elif a.cmd == "grammar":
        import grammar
        grammar.run(rdir, load_config(rdir))
    elif a.cmd == "cite":
        if a.action == "add":
            if not a.queries:
                sys.exit("usage: report cite add \"<title | DOI | arXiv id>\" ...")
            cite.cmd_add(rdir, a.queries)
        elif a.action == "verify":
            cite.cmd_verify(rdir, fix=a.fix)
        else:
            cite.cmd_review(rdir)
    elif a.cmd == "inspect":
        for level, msg in inspect_layout.inspect(rdir, load_config(rdir)):
            print(f"  {level}: {msg}")


if __name__ == "__main__":
    main()
