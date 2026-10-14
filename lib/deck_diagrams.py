"""Diagrams and charts drawn for slides, not pages.

Report diagrams are made for an A4 text column: tall, small text. On a 16:9
slide they shrink until nobody at the back can read them. Sources in
slides/diagrams/ are rendered with a slide style instead: left-to-right flow,
the deck's font at a large size, PNG at 3x (python-pptx can't place SVG; the
SVG is kept only to measure text size). slides/charts/*.py run like the
report's charts, with styles/slides.mplstyle available for big, wide plots.
"""
import re
import subprocess
import tempfile
from pathlib import Path

KIT = Path(__file__).resolve().parent.parent


def fc_file(family, style):
    r = subprocess.run(["fc-match", "-f", "%{file}", f"{family}:style={style}"], capture_output=True, text=True)
    p = r.stdout.strip()
    return p if r.returncode == 0 and p.endswith((".ttf", ".otf")) and Path(p).exists() else None


def slide_mermaid(src: str) -> str:
    """Top-down flowcharts become left-to-right: slides are wide, not tall."""
    return re.sub(r"^(\s*(?:flowchart|graph))\s+(TD|TB)\b", r"\1 LR", src, count=1, flags=re.M)


def slide_d2(src: str) -> str:
    return src if re.search(r"^direction\s*:", src, flags=re.M) else "direction: right\n" + src


def render(sdir: Path, cfg: dict, family="Calibri"):
    import report                       # run(), stale(), the kit PATH
    figs = sdir / "figures"
    figs.mkdir(exist_ok=True)
    n = 0
    mm_cfg = sdir / "_build" / "mermaid-slides.json"
    mm_cfg.parent.mkdir(exist_ok=True)
    mm_cfg.write_text('{"theme": "neutral", "htmlLabels": false, "flowchart": {"htmlLabels": false, '
                      '"curve": "basis", "nodeSpacing": 40, "rankSpacing": 50}, '
                      f'"themeVariables": {{"fontFamily": "{family}", "fontSize": "22px"}}}}')
    d2_fonts = []
    for flag, style in (("--font-regular", "Regular"), ("--font-bold", "Bold"), ("--font-italic", "Italic")):
        f = fc_file(family, style)
        if f:
            d2_fonts += [flag, f]
    for src in sorted((sdir / "diagrams").glob("*")) if (sdir / "diagrams").exists() else []:
        svg, png = figs / f"{src.stem}.svg", figs / f"{src.stem}.png"
        if not report.stale(src, [svg, png]):
            continue
        with tempfile.TemporaryDirectory() as tmp:
            if src.suffix == ".mmd":
                t = Path(tmp) / src.name
                t.write_text(slide_mermaid(src.read_text()))
                for out, extra in ((svg, []), (png, ["-s", "3"])):
                    report.run(["mmdc", "-i", t, "-o", out, "-b", "white", "-c", mm_cfg, *extra], quiet=True)
            elif src.suffix == ".d2":
                t = Path(tmp) / src.name
                t.write_text(slide_d2(src.read_text()))
                report.run(["d2", "--layout=elk", "--pad", "20", *d2_fonts, t, svg], quiet=True)
                report.run(["d2", "--layout=elk", "--pad", "20", "--scale", "3", *d2_fonts, t, png], quiet=True)
            else:
                continue
        n += 1
    for src in sorted((sdir / "charts").glob("*.py")) if (sdir / "charts").exists() else []:
        marker = figs / f".{src.stem}.done"
        if report.stale(src, [marker]):
            report.run([KIT / ".venv/bin/python", src], cwd=sdir, quiet=True)
            marker.touch()
            n += 1
    if n:
        report.say(f"rendered {n} slide diagram/chart source(s)")


def smallest_text_pt(img: Path, placed_width_pt: float):
    """Smallest text size (pt) a rendered diagram will have on the slide, from its sibling SVG."""
    svg = img.with_suffix(".svg")
    if not svg.exists():
        return None
    s = svg.read_text(errors="replace")
    m = re.search(r'viewBox="[-\d.]+\s+[-\d.]+\s+([\d.]+)\s+[\d.]+"', s) or re.search(r'width="([\d.]+)(px)?"', s)
    if not m:
        return None
    width = float(m.group(1))
    # CSS for things a still picture never shows (tooltips, error text) doesn't count
    s = re.sub(r"[^{}]*(?:[Tt]ooltip|[Ee]rror|TitleText)[^{}]*\{[^}]*\}", "", s)
    sizes = [float(x) for x in re.findall(r'font-size[:=]\s*"?([\d.]+)(?:px)?', s)]
    if not sizes or width <= 0:
        return None
    return min(sizes) * placed_width_pt / width
