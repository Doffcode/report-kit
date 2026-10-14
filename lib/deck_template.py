"""report slides template-extract college.pptx --name nitt

Turns a college/company .pptx into a deck template: templates/slides/nitt.pptx
plus templates/slides/nitt.yml, the slide-type → layout mapping it guessed.
Check the .yml once; fix a layout name there if a guess is wrong.

College templates are often hand-made: the logo pasted on every sample slide
instead of the master, one "Title and Content" layout for everything, banners
drawn on the sample slides. So:
  - a picture or shape that sits at the same spot on most sample slides is
    moved onto the master (all new slides then get it, and it's editable once);
  - slide types with no matching layout are listed; the builder places those
    with positioned shapes on the closest layout instead;
  - fonts the template uses but this PC lacks are reported (fitting needs them).
"""
import copy
import hashlib
import sys
from pathlib import Path

import yaml
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

KIT = Path(__file__).resolve().parent.parent
KINDS = {"title": "title slide", "section": "section divider", "content": "title + bullets",
         "two": "two columns", "title_only": "title only (pictures, tables, charts, results)", "blank": "blank"}


def shape_key(sh):
    """Same element at the same spot: position + kind + content (image bytes or XML text)."""
    if sh.is_placeholder:
        return None
    pos = (round(sh.left or 0, -4), round(sh.top or 0, -4), round(sh.width or 0, -4), round(sh.height or 0, -4))
    if sh.shape_type == 13:
        return ("pic", pos, hashlib.sha1(sh.image.blob).hexdigest())
    xml = copy.deepcopy(sh._element)
    for e in xml.iter(qn("p:cNvPr")):
        e.set("id", "0")
        e.set("name", "")
    from lxml import etree
    return ("shape", pos, hashlib.sha1(etree.tostring(xml)).hexdigest())


def repeated_elements(prs):
    """Elements on at least 60% of the sample slides (and ≥ 2 slides)."""
    slides = list(prs.slides)
    if len(slides) < 2:
        return []
    count, first = {}, {}
    for s in slides:
        for sh in s.shapes:
            k = shape_key(sh)
            if k is None:
                continue
            count[k] = count.get(k, 0) + 1
            first.setdefault(k, sh)
    need = max(2, int(len(slides) * 0.6 + 0.999))
    return [first[k] for k, c in count.items() if c >= need]


def move_to_master(prs, sh):
    master = prs.slide_master
    el = copy.deepcopy(sh._element)
    if sh.shape_type == 13:                       # picture: relate its image part to the master
        blip = el.find(".//" + qn("a:blip"))
        old = blip.get(qn("r:embed"))
        part = sh.part.related_part(old)
        blip.set(qn("r:embed"), master.part.relate_to(part, RT.IMAGE))
    tree = master.shapes._spTree
    ids = [int(e.get("id")) for e in tree.iter(qn("p:cNvPr"))]
    el.find(".//" + qn("p:cNvPr")).set("id", str(max(ids + [1]) + 1))
    tree.append(el)


def main(src: Path, name: str):
    if not src.exists():
        sys.exit(f"{src}: not found")
    import deck
    problems = []
    prs = Presentation(str(src))
    moved = []
    already = {shape_key(sh) for sh in prs.slide_master.shapes}
    for sh in repeated_elements(prs):
        if shape_key(sh) not in already:
            move_to_master(prs, sh)
            moved.append(f"{'picture' if sh.shape_type == 13 else 'shape'} “{sh.name}”")
    n_samples = len(prs.slides._sldIdLst)
    out_dir = KIT / "templates/slides"
    dest = out_dir / f"{name}.pptx"
    if dest.exists() and dest.resolve() != src.resolve():
        print(f"overwriting {dest}")
    prs.save(dest)

    tpl = deck.Template(dest, problems)       # same layout guessing the builder uses
    found = {k: tpl._find(k) for k in KINDS}
    tpl.fonts.file("+mn-lt")                    # missing fonts are reported into problems
    tpl.fonts.file("+mj-lt", "b")
    mapping = {
        "source": src.name,
        "slide_size_in": [round(tpl.W / 914400, 2), round(tpl.H / 914400, 2)],
        "fonts": {"titles": tpl.fonts.major, "body": tpl.fonts.minor},
        "layouts": {k: (v.name if v is not None else None) for k, v in found.items()},
        "available_layouts": [lay.name for lay in tpl.prs.slide_layouts],
        "moved_to_master": moved,
    }
    note = ("# Made by `report slides template-extract`. Check once: every slide type should name a\n"
            "# layout from available_layouts (null = the builder places it on the closest layout).\n")
    (out_dir / f"{name}.yml").write_text(note + yaml.safe_dump(mapping, sort_keys=False, allow_unicode=True))

    print(f"\n  template {name}: {dest.relative_to(KIT)} + {name}.yml")
    print(f"  slide size {mapping['slide_size_in'][0]} × {mapping['slide_size_in'][1]} in"
          + ("" if abs(tpl.W / tpl.H - 16 / 9) < 0.02 else "  (not 16:9: projectors today are wide)"))
    for k, desc in KINDS.items():
        lay = found[k]
        print(f"  {'✓' if lay is not None else '·'} {desc:<46} → {lay.name if lay is not None else 'none: closest layout'}")
    for m in moved:
        print(f"  ✓ moved {m} from the sample slides onto the master")
    if n_samples:
        print(f"  · its {n_samples} sample slide(s) are dropped when a deck is built")
    for level, msg in problems:
        print(f"  {'✗' if level == 'FAIL' else '!'} {msg}")
    print(f"\n  use it: template: {name}   (in slides/slides.yml)")
