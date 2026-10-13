# Report Kit

Turns a project into a college-format report (PDF + Word) with one command.
You write plain content; the build does every bit of layout, so nothing can
shift, break or go stale when you edit.

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

## Inside

Quarto → Typst for the PDF (`templates/report/`), Quarto → Word with
`lib/docx_fix.py` + `lib/lo_refresh.py` for the DOCX. Both read the same
filled-in format profile (`lib/profile.py`), so they can't drift apart. Tools are installed
user-local under `tools/`; versions are pinned in `versions.txt`.
