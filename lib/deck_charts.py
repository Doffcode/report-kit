"""Charts on slides from a CSV: ![caption](charts/acc.csv){type=line x="Round" y="Accuracy"}

CSV: first column = x / categories, every other column = one series. Columns
named <series>_std or <series>_err give error bands/bars.

Plain bar/line charts become native PowerPoint charts (the numbers stay
editable in PowerPoint: right-click → Edit Data). python-pptx can't draw error
bars or log axes, so those charts are drawn with matplotlib as a picture.
"""
import csv
from pathlib import Path

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Pt

KIT = Path(__file__).resolve().parent.parent
ERR = ("_std", "_err")
TYPES = {"line": XL_CHART_TYPE.LINE_MARKERS, "bar": XL_CHART_TYPE.COLUMN_CLUSTERED,
         "hbar": XL_CHART_TYPE.BAR_CLUSTERED}


def read_csv(path: Path):
    with open(path, newline="") as fh:
        rows = [r for r in csv.reader(fh) if r and any(c.strip() for c in r)]
    head, body = rows[0], rows[1:]
    cols = {h.strip(): [r[k].strip() if k < len(r) else "" for r in body] for k, h in enumerate(head)}
    x = head[0].strip()
    series = [h.strip() for h in head[1:] if not h.strip().endswith(ERR)]
    errs = {s: next((cols[s + e] for e in ERR if s + e in cols), None) for s in series}
    num = lambda v: float(v) if v not in ("", "NA", "nan") else None  # noqa: E731
    return {"x_name": x, "x": cols[x], "series": {s: [num(v) for v in cols[s]] for s in series},
            "err": {s: [num(v) or 0 for v in e] if e else None for s, e in errs.items()}}


def needs_picture(data, attrs):
    return any(data["err"].values()) or attrs.get("log") or attrs.get("picture")


def native(slide, box, data, attrs, ink=None):
    l, t, w, h = box
    cd = CategoryChartData()
    cd.categories = data["x"]
    for name, vals in data["series"].items():
        cd.add_series(name, vals)
    kind = attrs.get("type", "line").strip('"')
    chart = slide.shapes.add_chart(TYPES.get(kind, TYPES["line"]), l, t, w, h, cd).chart
    chart.font.size = Pt(16)
    if ink:                             # a dark slide: light labels and grid
        chart.font.color.rgb = RGBColor.from_string(ink)
    chart.has_title = False             # the slide title says what the chart shows
    chart.has_legend = len(data["series"]) > 1
    if chart.has_legend:
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
    va = chart.value_axis
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = RGBColor(0xD9, 0xD9, 0xD9) if not ink else RGBColor(0x5A, 0x5A, 0x5A)
    va.format.line.fill.background()
    for key, attr in (("ymin", "minimum_scale"), ("ymax", "maximum_scale")):
        if attrs.get(key):
            setattr(va, attr, float(attrs[key].strip('"')))
    for axis, key in ((chart.category_axis, "x"), (va, "y")):
        label = (attrs.get(key) or (data["x_name"] if key == "x" else "")).strip('"')
        if label:
            axis.has_title = True
            axis.axis_title.text_frame.text = label
            axis.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(16)
            axis.axis_title.text_frame.paragraphs[0].runs[0].font.bold = False
            if ink:
                axis.axis_title.text_frame.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string(ink)
    if kind == "line":
        for s in chart.plots[0].series:
            s.smooth = False
            s.format.line.width = Pt(2.75)
    return chart


def matplotlib_png(data, attrs, out_dir: Path, stem, family="Calibri", aspect=16 / 7, colours=None, ink=None,
                   bg="FFFFFF"):
    """colours: the theme's accents (series colours); ink: text colour on a dark slide; bg: slide colour."""
    import os
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    out_dir.mkdir(parents=True, exist_ok=True)
    from matplotlib import font_manager
    import deck_diagrams
    names = []
    for style in ("Regular", "Bold"):        # the font fontconfig uses for the deck (Calibri → Carlito)
        f = deck_diagrams.fc_file(family, style)
        if f:
            font_manager.fontManager.addfont(f)
            names.append(font_manager.FontProperties(fname=f).get_name())
    plt.rcdefaults()                         # settings from an earlier (dark) chart must not carry over
    plt.style.use(str(KIT / "styles/slides.mplstyle"))
    plt.rcParams["font.sans-serif"] = names[:1] + plt.rcParams["font.sans-serif"]
    if colours:
        plt.rcParams["axes.prop_cycle"] = plt.cycler(color=["#" + c for c in colours])
    plt.rcParams["figure.facecolor"] = plt.rcParams["axes.facecolor"] = plt.rcParams["savefig.facecolor"] = "#" + bg
    if ink:
        for k in ("text.color", "axes.labelcolor", "xtick.color", "ytick.color", "axes.edgecolor"):
            plt.rcParams[k] = "#" + ink
        plt.rcParams["grid.color"] = "#5A5A5A"
        plt.rcParams["legend.facecolor"] = "#" + bg
    fig, ax = plt.subplots(figsize=(12, 12 / max(aspect, 1.2)))
    kind = attrs.get("type", "line").strip('"')
    try:
        xs = [float(v) for v in data["x"]]
    except ValueError:
        xs = list(range(len(data["x"])))
        ax.set_xticks(xs, data["x"])
    names = list(data["series"])
    width = 0.8 / max(len(names), 1)
    for k, name in enumerate(names):
        ys, err = data["series"][name], data["err"][name]
        if kind == "bar":
            pos = [x + (k - (len(names) - 1) / 2) * width for x in xs]
            ax.bar(pos, ys, width, yerr=err, capsize=5, label=name)
        else:
            ax.plot(xs, ys, marker="o", label=name)
            if err:
                ax.fill_between(xs, [y - e for y, e in zip(ys, err)], [y + e for y, e in zip(ys, err)], alpha=0.2)
    if attrs.get("log"):
        ax.set_yscale("log")
    ax.set_xlabel((attrs.get("x") or data["x_name"]).strip('"'))
    if attrs.get("y"):
        ax.set_ylabel(attrs["y"].strip('"'))
    if attrs.get("ymin") or attrs.get("ymax"):
        ax.set_ylim(float(attrs["ymin"]) if attrs.get("ymin") else None,
                    float(attrs["ymax"]) if attrs.get("ymax") else None)
    if len(names) > 1:
        ax.legend()
    png = out_dir / f"chart-{stem}.png"
    fig.savefig(png)
    plt.close(fig)
    os.utime(png)
    return png
