# Report Kit

Turns a project into a college-format report (PDF + Word) and a slide deck
(editable PowerPoint) with one command each. You write plain content; the build
does every bit of layout, so nothing can shift, break or go stale when you edit.

## Install (any Ubuntu PC, no sudo)

    # 1. sign in to GitHub (the repo is private); portable gh, no sudo
    curl -fsSL https://github.com/cli/cli/releases/download/v2.102.0/gh_2.102.0_linux_amd64.tar.gz | tar -xz -C ~
    ~/gh_2.102.0_linux_amd64/bin/gh auth login        # choose GitHub.com → HTTPS → login with a web browser
    ~/gh_2.102.0_linux_amd64/bin/gh repo clone Doffcode/report-kit ~/report-kit

    # 2. install everything
    ~/report-kit/install.sh

It downloads every tool into the kit folder (about 2 GB), installs the fonts,
the `report` command and the Claude Code skill, then builds a demo report to
prove it works. Update later with `cd ~/report-kit && git pull && ./install.sh`.
On a shared lab PC, sign out when you leave: `~/gh_2.102.0_linux_amd64/bin/gh auth logout`
(and `claude /logout` if you signed in to Claude Code there).

## Use it

    cd ~/Projects/my-project
    report init               # once: creates report/
    report build              # → report/out/report.pdf + report.docx
    report build --final      # before submitting: gaps become hard failures

Or in Claude Code: "make a report for this project" (the `report` skill).

Fill `~/report-kit/me.yml` once (name, roll no, guide, college) and put your
college logo at `~/report-kit/assets/logo.png`.

## Formats

    report init --profile thesis        # project (default) | thesis | internship | seminar

`report.yml` → `profile:` picks the format; `format:` overrides any setting of it
(spacing, margins, fonts, chapter style, caption labels, cover lines, certificate
wording ...). Formats live in `profiles/*.yml`; `base.yml` documents every key.
Your college's own format:

    report style-extract sample.pdf --name nitt-mtech   # measures a sample report → profiles/nitt-mtech.yml

It measures paper, margins, font, spacing, chapter style, captions, page numbers,
front-page order, the cover and the certificate/declaration wording (with the
sample's names swapped for placeholders). Tested round trip: a report rebuilt
from an extracted profile matched the sample on 13 of 14 pages, the cover's
spacing being the approximate part. Check a new profile against the sample once.

One-shot mode: give Claude everything at once; it writes every guess to
`report/assumptions.md`, highlights gaps in yellow, and `--final` refuses to
pass until you have confirmed them.

## Before submitting

    report grammar         # local LanguageTool pass (nothing leaves the machine) → _build/grammar.md
    report build --final   # every gap becomes a failure

Every build also flags sentences that copy 8+ words in a row from, or closely
paraphrase, a cited paper's abstract or anything in `sources/` (similarity
checkers like Turnitin catch both).

## References

    report cite add "Deep Learning with Differential Privacy"   # by title
    report cite add 10.1145/2976749.2978318 arXiv:1602.05629    # by DOI / arXiv id
    report cite verify [--fix]   # check every refs.bib entry against Crossref, DBLP, OpenAlex, arXiv
    report cite review           # each citing sentence next to the paper's abstract

References are fetched, never typed from memory. Each entry is marked
`verified = {crossref 2026-10-13}` (or `openalex`, `dblp`, ...). Papers whose
venue can't be confirmed automatically (no DOI, e.g. MLSys) or that aren't
found at all stay flagged until you check them once and set
`verified = {manual <date>}`. The published version wins over an arXiv
preprint. `citation_style: ieee` or `apa` in report.yml sets both PDF and Word.
Set SEMANTIC_SCHOLAR_API_KEY to add Semantic Scholar as an abstract source.

## What every build checks

- Unresolved cross-references or citations, missing images, Type 3 / fallback fonts
- Leftover `[[TODO]]`s, figures never mentioned in the text
- Numbers in the chapters that aren't in `facts.md`
- Abbreviations not in the glossary or not spelled out at first use
- Citations: missing from refs.bib, not verified, venue unconfirmed, or the
  cited paper's abstract shares little with the sentence citing it
- Text copied from, or closely paraphrasing, a cited abstract or a file in sources/
- Cover/front-page details missing from me.yml (shown as yellow TODOs), and
  one-shot assumptions not yet confirmed
- Layout, on both the PDF and a rendering of the Word file: text past the margins,
  half-empty pages mid-chapter, headings stranded at a page bottom, captions
  separated from their figures, text too small to read (shrunken diagrams)
- What moved since the last build (e.g. "Chapter 5 p.33 → p.34")

`report/_build/pages.png` shows every page at a glance.

## Slides

```bash
report slides init                  # creates slides/ next to report/
report slides init --from-report    # … or drafts slides.md from the report's sections
report slides build                 # → slides/out/slides.pptx + slides.pdf, _build/slides.png
report slides build --final         # before the talk: TODOs and missing speaker notes fail
```

Or in Claude Code: "make the slides for this project". One `## Title` per slide in
`slides/slides.md`; write the title as the slide's point ("FedProx beats FedAvg on
non-IID clients"), keep details in `::: notes`. Pictures, `| tables |`, CSV charts
(`![](data/acc.csv){type=line}`), big-number `{.results}` cards, `$$ equations $$`,
`[@citations]` and `{.backup}` slides for viva questions are all built for you.

Unlike Markdown-to-slides tools that export pictures of slides, every title, bullet,
table and simple chart is a real PowerPoint object: change anything in PowerPoint,
Google Slides or Canva. Text is measured with the template's real fonts and sized to
fit (never shrunk by PowerPoint afterwards), pictures are never cropped, and text
that still doesn't fit is split into "(cont.)" slides and reported.

**Look.** A theme is a style, a palette and a font pair, set in `slides/slides.yml`:

```yaml
theme:
  style: band        # clean | band | sidebar | minimal | bold | soft | geometric
  palette: teal      # ocean royal teal forest mint sunset crimson berry plum slate earth gold mono
  # primary: "#7B2CBF"   # or your own colour: the rest of the palette is derived from it
  fonts: classic     # classic | serif-titles | formal | plain | times
cover: college       # first slide: logo + college name, title, "Presented by" / "Under the guidance of"
# cover: first-page.pptx   # or your college's own first slide, with {{title}} {{name}} {{guide}} {{date}} in it
# logo_on_slides: true     # small logo on every slide
# team: [{name: "A", roll: "1"}, {name: "B", roll: "2"}]
```

`report slides themes` draws every style and palette (`_build/themes.png`, `palettes.png`).

**Designs from the content.** A slide's content picks its design, and the build lists
every choice with the reason: a sentence followed by points → the statement beside its
points; a numbered list → big numerals (or process arrows when the title says how/steps);
two lists under `### Existing` / `### Proposed {.win}` → side by side with "vs", the better
one outlined; `### Advantages` / `### Limitations` → ✓ / ✗ columns; three or more `###`
groups → columns; a picture with 2–4 short notes → the picture with numbered notes; a
chart and one sentence → the chart with a takeaway box; `.results` numbers with a chart →
numbers beside the chart; `- **Term**: meaning` alone → a definition; "Conclusion" →
numbered takeaways; "Future work" → a roadmap. Sections get an outline slide and a
"you are here" list on each divider; a table's Gap/Limitation column and the "ours" row
are highlighted. The accent colour marks only what matters on each slide. Icons
(`- :lock: **Privacy**: …`, any name from lucide.dev/icons) are added automatically only
when every item in a group has an obvious one. Every design can also be asked for:
`{.lead}` `{.lead-cards}` `{.numbered}` `{.checklist}` `{.compare}` `{.proscons}` `{.columns}`
`{.callouts}` `{.insight}` `{.definition}` `{.takeaways}` `{.roadmap}`.

**Slide designs.** Besides bullets, pictures, tables and charts: `{.cards}` (items as
cards: `- **Label**: text`), `{.steps}` (numbered arrows), `{.timeline}` (`- Jan :: Survey`),
`{.statement}` (one big message), `{.quote}`, `{.image-left}` / `{.image-right}`. Any slide
can have its own background: `{.dark}`, `{.accent}`, `{.light}`, `{color="#E63946"}` or
`{bg="photo.jpg"}`. With `variety: auto` (the default) plain slides that read better
another way are drawn that way: labelled bullets → cards, a short numbered list → steps,
a single sentence → statement, picture + text → picture in its own half, alternating
sides. The build lists what it changed; `{.bullets}` or `variety: off` keeps plain bullets.

A college/company template instead of a kit theme:

```bash
report slides template-extract college.pptx --name nitt   # → templates/slides/nitt.pptx + nitt.yml
```

then `template: nitt` in slides.yml. Check `nitt.yml` once (which layout is used for
which slide type). Logos and banners pasted on every sample slide are moved to the
master; slide types the template has no layout for are placed on the closest one, kept
clear of its logo. Open the deck once in PowerPoint (or PowerPoint for the web) before
the talk: the preview is LibreOffice's rendering, which is close but not identical.

## Inside

Quarto → Typst for the PDF (`templates/report/`), Quarto → Word with
`lib/docx_fix.py` + `lib/lo_refresh.py` for the DOCX. Both read the same
filled-in format profile (`lib/profile.py`), so they can't drift apart. Tools are installed
user-local under `tools/`; versions are pinned in `versions.txt`.

Slides: `lib/deck.py` (python-pptx into the template's layouts, font-metric fitting),
`lib/deck_charts.py`, `lib/deck_diagrams.py`, `lib/deck_check.py`, `lib/deck_template.py`;
kit themes are made at build time by `lib/deck_theme.py`.
