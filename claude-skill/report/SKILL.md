---
name: report
description: Make or update a project report (college project report, thesis, internship/seminar report) as PDF + DOCX, or a slide deck (review/viva, seminar or any talk) as an editable PPTX, with the Report Kit at ~/report-kit. Use when the user asks to make, write, build, fix or change a report or presentation/slides/PPT for a project or paper, or types /report.
---

# Report Kit workflow

**The user's whole job is one sentence ("make the report for this project").** Do every
step yourself, including the ones that look like "the user's data": run the project's
tests, measure missing results by running the project's own code (simulate clients/rounds
with its real training + aggregation functions, keep the script in report/sources/experiments/),
confirm references, draft the acknowledgement from me.yml (no pronouns for people whose
pronouns you don't know). Only hand back what truly can't be produced: signatures, a company's
certificate scan, facts that exist nowhere in the sources. Never hand the user a to-do list
for things you could do.

The kit lives in `~/report-kit` (README there). One command does all layout: `report build`.
Never format anything by hand and never edit `out/report.docx`: every build overwrites it.
All changes go into the markdown sources, `report.yml`, or the kit templates.

## Folder (inside the project)

```
report/
  report.yml      title, profile, date, target_pages, chapters order, not_abbreviations,
                  citation_style: ieee | apa, internship: {...}, format: {overrides of any profile key}
  assumptions.md  one-shot mode: every guess made without asking (user deletes lines once confirmed)
  me.yml          optional; overrides ~/report-kit/me.yml (name, roll, guide, college, logo)
  sources/        anything the user gives: PDFs, docs, AI-chat exports, base paper, notes
  facts.md        every number/claim the report may use, each with [source: file:line]
  glossary.md     "- ABBR: Expansion" lines -> List of Abbreviations + first-use check
  refs.bib        BibTeX via `report cite`; cite with [@key]; every entry carries verified = {...}
  refs/           abstracts.json (claim check), generated typst.bib + style .csl
  front/          acknowledgement, abstract ({.unnumbered} H1 each)
  chapters/       NN-name.md, one H1 per file = one chapter
  diagrams/       *.mmd (Mermaid) and *.d2 -> figures/name.svg + .png automatically
  charts/         *.py matplotlib scripts that READ real data and write figures/name.svg+png
  out/            report.pdf, report.docx (generated)
  _build/         logs, pages.png overview, docx-preview.pdf (generated)
```

## Steps

1. **Intake.** `report init --profile project|thesis|internship|seminar` (creates `report/`). Ask only what can't be found: format, page target, anything the college mandates. Internship: fill `internship:` in report.yml (company, place, period, mentor, mentor_title) and ask for the certificate scan → `assets/internship-certificate.png`.
   **One-shot mode** (user gives everything at once, or says "just make it"): ask nothing. Pick the profile from context, run the whole pipeline to a draft build, and write every guess to `report/assumptions.md` as bullets ("- Assumed profile: thesis, because ..."). Gaps stay as `[[TODO: ...]]` (highlighted yellow in both outputs). The checker warns on assumptions and TODOs, and a `--final` build refuses them until the user confirms.
2. **Sources.** Read the repo (README, docs, code, notebooks incl. outputs, logs, tests, package files) and everything in `sources/`. Convert binary docs with `~/report-kit/.venv/bin/markitdown file` (PDF, DOCX, PPTX, XLSX, HTML).
3. **facts.md.** Terse bullets, each ending `[source: path:line]`. Copy numbers exactly. Missing -> `[[TODO: ...]]`. Read the code, not just the docs: where docs claim more than the code does, write a `**CODE ≠ DOCS:**` line and make the report follow the code. Flag suspicious results (e.g. R² = 1.00 everywhere -> possible leakage).
4. **Outline.** Propose chapters from what the sources actually contain (the standard list is a guideline: intro, literature, requirements, design, methodology, implementation, testing, results, conclusion). Get one approval unless in one-shot mode.
5. **Write** chapters (see rules) + diagrams + charts + glossary. References: never type BibTeX from memory.
   `report cite add "<title | DOI | arXiv id>"` fetches each one (prints the key to cite). Existing
   refs.bib: `report cite verify` (add `--fix` to replace wrong fields with the real record).
   Then `report cite review` and READ `_build/citations.md`: judge for every ⚠ whether the paper
   really supports the sentence (the overlap score is only a hint); fix the sentence or the citation.
   "venue unconfirmed" / "not found" entries: check them yourself (official proceedings pages,
   publisher sites); if confirmed, extend lib/cite.py so the kit confirms that venue automatically
   next time (see PROCEEDINGS). Use `verified = {manual ...}` only for what no source can confirm,
   and say so in the summary.
6. **Build** `report build` from the project or report folder. Read every line of the summary and fix FAILs and WARNs at the source. Look at `_build/pages.png` yourself.
7. **Revise** on plain-English requests by editing sources and rebuilding. Never touch the DOCX.
8. **Polish.** `report grammar` → read `_build/grammar.md`, apply the real fixes (it misses some agreement errors; read the text too). Similarity warnings (“copies 8+ words…”, “closely paraphrases…”) must be rewritten in new words, not lightly reworded.
9. **Submit** with `report build --final`: TODOs, untraceable numbers, unreferenced figures, unconfirmed assumptions/references, copied text and page-limit misses become failures.

## Writing rules (the checker enforces most)

- Every number in a chapter must appear in facts.md. Small integers <= 10 are exempt.
- Expand each abbreviation at first use in body text: "Federated Learning (FL)"; add it to glossary.md. Brand names go in `not_abbreviations`.
- Every figure/table/algorithm is referenced in the text (`@fig-x`, `@tbl-x`, `@alg-x`, `@eq-x`).
- Figures: `![Caption.](figures/name){#fig-name width=85%}` with NO file extension.
- Tables: pipe tables; grid tables only for merged cells. Caption line `: Caption. {#tbl-name}` after the table.
- Equations: `$$ ... $$ {#eq-name}`. Norms as `\lVert x \rVert`.
- Algorithms: `::: {#alg-name}` numbered list of steps, blank line, caption sentence as LAST paragraph, `:::`.
- Citations `[@key]` only for entries in refs.bib. Never invent a reference; unknown details stay out.
- Diagrams must suit a portrait page. Architecture in D2 with `grid-columns: 3` + containers `grid-rows: N` (ELK ignores nested `direction`). Flows in Mermaid `flowchart TD` or `sequenceDiagram`. If the inspector reports text below 6pt, redraw narrower.
- Charts read their numbers from the project's data/notebook files; never retype numbers.
- Plain, specific prose; no filler. Tell the user to rewrite intro, discussion and conclusion in their own voice (similarity/AI checkers), and respect any project note asking them to write prose themselves.

## Formats (profiles)

`~/report-kit/profiles/*.yml`: base (all keys, documented), project, thesis, internship, seminar.
A profile sets paper, margins, font, spacing, chapter style (block/inline), heading sizes, caption
labels, front lists, cover lines and front pages (certificate/declaration wording, signatures) with
{placeholders} from me.yml/report.yml. Change one report: `format:` in report.yml. A college's own
format: copy a profile to profiles/<college>.yml with `extends: project`, or derive it from a sample
PDF with `report style-extract sample.pdf --name <college>`. PDF and DOCX both read the same profile.

## Layout guarantees (already in the kit; don't re-implement)

PDF (master): Typst template `~/report-kit/templates/report/`. Cover/certificate/declaration from me.yml, roman then arabic page numbers, CHAPTER N headings on new pages, per-chapter numbering (Fig. 4.2, (4.1)), floats that never leave half-empty pages or sit above a chapter title, sticky headings, images capped at 70% page height.
DOCX: reference.docx + `lib/docx_fix.py` (front pages, sections, captions, borders, keep-together, image fit, XML repair) + `lib/lo_refresh.py` (fills TOC/LOF/LOT via LibreOffice).
If a layout bug appears, fix it in the kit (template, docx_fix, filters) so every report benefits, then add a check to `lib/inspect_layout.py` or `lib/check.py`.

## Slides (editable PPTX)

Same idea as the report: Claude writes plain `slides/slides.md`, `report slides build` does all
layout into the layouts of a real .pptx, so every title, bullet, table and simple chart stays a
native PowerPoint object the user can edit (PowerPoint, Google Slides, Canva import).

```
slides/                (next to report/; report figures, refs.bib, facts.md, glossary.md are reused)
  slides.yml   title, subtitle, template: kit | <name> | college.pptx, talk_minutes, footer, qa
  slides.md    "# Part" = section divider, "## Claim as title" = slide, ::: notes ::: = speaker notes
  diagrams/    slide versions (*.mmd/*.d2): rendered left-to-right, big font. Don't reuse report diagrams as is.
  charts/      *.py with styles/slides.mplstyle (only when a CSV chart can't express it)
  data/        CSVs for charts
  out/         slides.pptx + slides.pdf;  _build/slides.png = every slide at a glance
```

**Steps.** `report slides init` (or `--from-report` to draft from the report's sections: bullets are
first sentences, full paragraphs go to notes, titles marked [[TODO]]). Rewrite it into the talk:
3–4 points a slide, the rest in notes. `report slides build`, read every line, open
`_build/slides.png` and LOOK at it; fix at the source and rebuild. Before presenting: `--final`
(TODOs, missing notes, unverified references fail).

**slides.md syntax.** Type is inferred or forced with `{.results} {.two-col} {.image} {.references} {.qa} {.backup}`.
- bullets → bullet slide; text then `|||` then text/picture → two columns; text + one picture → side by side (wide pictures and equations go below the text automatically)
- `![Caption.](figures/x)` picture, fitted without cropping; `figures/` of slides/ and report/ are searched
- `![Caption.](data/acc.csv){type=line|bar|hbar x="Round" y="Accuracy" ymin=0.5}` → native editable chart; columns `<series>_std`/`_err` or `log=y` → matplotlib picture (python-pptx can't draw those)
- `| pipe | table |` → native table, sized to its content
- `.results`: `- 0.92 :: FedProx accuracy` → big-number cards
- `$$ LaTeX $$` → equation picture; short inline `$\mu$` → text (μ)
- `[@key]` → [n], numbered in order; the `{.references}` slide lists exactly what the slides cite
- `{.backup}` slides go after the automatic Thank-you/Questions slide, for likely viva questions

**Content rules (checked).** Titles state the claim ("FedProx beats FedAvg on non-IID clients"), not the
section name ("Results"). ≤ 6 bullets and ≤ 40 words a slide. Numbers from facts.md, abbreviations
spelled out once, about one content slide per minute of `talk_minutes`, notes on every content slide.

**Look (make every deck look different, not one default).** In slides.yml, `theme:` = style
(clean | band | sidebar | minimal | bold | soft | geometric) + palette (ocean royal teal forest mint sunset
crimson berry plum slate earth gold mono, or `primary: "#hex"`) + fonts (classic | serif-titles | formal |
plain | times); short form `theme: band teal`. Ask the user which they like, or show them
`report slides themes` (→ _build/themes.png, palettes.png). `cover: college` = logo + college name + title +
"Presented by" / "Under the guidance of" (from me.yml; several students: `team:`). `cover: x.pptx` copies
the college's own first slide and fills {{title}} {{name}} {{roll}} {{team}} {{guide}} {{guide_title}}
{{department}} {{college}} {{date}}. `logo_on_slides: true` puts a small logo on every slide.

**Designs follow content (the user's taste: clean, intentional, no filler decoration).** Write the
content in the shape that says what it is, and the kit picks the design (build output lists each choice):
statement paragraph + 2–4 points → lead; ordered list → numerals (process title → steps); `### A` /
`### B {.win}` → compare with vs and the winner outlined; `### Advantages` / `### Limitations` → ✓/✗;
3+ `###` groups → columns; picture + 2–4 short bullets → callouts; chart + one sentence → insight
(takeaway box); `.results` + chart → numbers beside chart; single `**Term**: meaning` → definition;
"Conclusion" → takeaways; "Future work" → roadmap. Agenda + "you are here" dividers come from `#`
sections (`agenda: false` to turn off). Literature tables: name a column "Gap" and our row "This work".
Icons `:name:` (Lucide) only where they clearly match; auto only if every item gets one. Explicit classes:
lead, lead-cards, numbered, checklist, compare, proscons, columns, callouts, insight, definition,
takeaways, roadmap.

**Slide designs.** `{.cards}` (`- **Label**: text`), `{.steps}` (process), `{.timeline}` (`- Jan :: text`),
`{.statement}` (title = the message, optional line under it), `{.quote}` (paragraph + `— who`),
`{.image-left}`/`{.image-right}`. Per-slide background: `{.dark}` `{.accent}` `{.light}` `{color="#hex"}`
`{bg="photo.jpg"}`. `variety: auto` (default) turns plain slides into cards/steps/statement/split picture
where they fit and reports it; use the designs on purpose so the deck has rhythm (not 15 bullet slides).

**Templates.** A college template (when there is a full one): `report slides template-extract college.pptx
--name <college>` → templates/slides/<college>.pptx + <college>.yml (slide type → layout; READ it once and
fix wrong guesses), then `template: <college>`. It moves logos/banners pasted on every sample slide onto the
master and reports missing fonts. Missing slide types fall back to the closest layout; text is kept clear of
the template's logo and banner.

**Layout guarantees (don't re-implement).** Text is measured with the real font (Calibri ↔ Carlito,
Cambria ↔ Caladea, Arial ↔ Liberation Sans) and sized to fit with a 10% margin, autofit off; too much
text is split into "(cont.)" slides with a warning; a font that isn't installed fails the build;
pictures never crop; diagrams whose text would be < 14 pt are flagged; the rendered PDF is checked for
text over the footer or off the slide. The LibreOffice preview is close, not exact: before the real
talk, ask the user to open the deck once in PowerPoint (or PowerPoint for the web / Google Slides).
