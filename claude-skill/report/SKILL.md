---
name: report
description: Make or update a project report (college project report, thesis, internship/seminar report) as PDF + DOCX with the Report Kit at ~/report-kit. Use when the user asks to make, write, build, fix or change a report for a project, or types /report.
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
