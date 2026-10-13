// ─────────────────────────────────────────────────────────────
// Report Kit: report layout.
// Layout rules live here; everything format-specific (margins, spacing,
// cover, certificate wording...) comes from `profile`, generated from
// ~/report-kit/profiles/*.yml. Chapters are plain markdown; nothing is
// positioned by hand, so edits can never break pagination rules.
// ─────────────────────────────────────────────────────────────

#import "kit-profile.typ": profile

#let kit-font = (profile.font.family, profile.font.fallback)
#let fig-kinds = ("quarto-float-fig", "quarto-float-tbl", "quarto-float-alg")
#let md(s) = eval(s, mode: "markup")      // profile text: *bold*, _italic_, #highlight[...]

// Chapter-relative number, e.g. "4.2". Front matter (chapter 0) gives plain "2".
#let chap-num(n) = {
  let ch = counter(heading).get().first()
  if ch == 0 { str(n) } else { str(ch) + "." + str(n) }
}

#let num-style(kind) = if kind == "roman" { "i" } else { "1" }

// Switch from front-matter to body page numbers. Called once, right before
// Chapter 1, by the generated report.qmd.
#let mainmatter(body) = {
  pagebreak(weak: true)
  set page(numbering: num-style(profile.page.numbers.body))
  counter(page).update(1)
  body
}

// Table of contents, list of figures, list of tables, in the profile's order.
#let frontlists() = {
  for item in profile.front_lists {
    if item == "toc" {
      outline(title: [Table of Contents], depth: 3, indent: auto)
    } else if item == "lof" {
      context if query(figure.where(kind: "quarto-float-fig")).len() > 0 {
        outline(title: [List of Figures], target: figure.where(kind: "quarto-float-fig"))
      }
    } else if item == "lot" {
      context if query(figure.where(kind: "quarto-float-tbl")).len() > 0 {
        outline(title: [List of Tables], target: figure.where(kind: "quarto-float-tbl"))
      }
    }
  }
}

// Quarto passes this name explicitly to every numbered equation; shadow it.
#let equation-numbering = n => "(" + chap-num(n) + ")"

#let sign-block(s) = align(center, block(width: 45%)[
  #set par(justify: false)
  #v(2.2cm)
  #line(length: 100%, stroke: 0.5pt)
  #md(s.name) \
  #text(size: 11pt, md(s.role))
])

#let report(doc) = {
  let P = profile
  let size = P.font.size * 1pt
  set document(title: P.doc_title)
  set page(
    paper: P.page.paper,
    margin: P.page.margin,
    numbering: none,
    footer: context {
      let n = page.numbering
      let where = if P.page.numbers.align == "right" { right } else { center }
      if n != none { align(where, text(size: 11pt, counter(page).display(n))) }
    },
  )
  // Text edges at ascender/descender so `leading` maps to Word's line spacing.
  set text(font: kit-font, size: size, lang: "en", region: "IN",
           top-edge: "ascender", bottom-edge: "descender")
  set par(justify: P.justify, leading: P.leading,
          spacing: P.leading + P.paragraph_after * 1pt, first-line-indent: 0em)
  show math.equation: set text(font: ("New Computer Modern Math",), top-edge: "cap-height", bottom-edge: "baseline")
  show raw: set text(font: ("DejaVu Sans Mono",), size: 10pt)

  // ── Headings ───────────────────────────────────────────────
  set heading(numbering: "1.1")
  show heading: set text(font: kit-font)
  show heading: set block(sticky: true)          // never stranded at a page bottom
  show heading.where(level: 1): it => {
    if P.chapter.new_page or it.numbering == none { pagebreak(weak: true) }
    // Restart figure/table/equation numbering in every chapter.
    for k in fig-kinds { counter(figure.where(kind: k)).update(0) }
    counter(math.equation).update(0)
    let title = if P.chapter.upper { upper(it.body) } else { it.body }
    if P.chapter.style == "inline" and it.numbering != none {
      block(above: 1.2cm, below: 0.9cm, text(size: P.headings.h1 * 1pt, weight: "bold")[
        #counter(heading).display("1")#h(0.8em)#title])
    } else {
      set align(center)
      v(1.2cm)
      if it.numbering != none {
        text(size: (P.headings.h1 - 2) * 1pt, weight: "bold")[#P.chapter.label #counter(heading).display("1")]
        v(0.3cm)
      }
      text(size: P.headings.h1 * 1pt, weight: "bold", title)
      v(0.9cm)
    }
  }
  show heading.where(level: 2): set text(size: P.headings.h2 * 1pt, weight: "bold")
  show heading.where(level: 2): set block(above: 1.6em, below: 1em)
  show heading.where(level: 3): set text(size: P.headings.h3 * 1pt, weight: "bold")
  show heading.where(level: 3): set block(above: 1.4em, below: 0.9em)

  // ── Figures, tables, equations: numbered per chapter (Fig. 4.2) ──
  show figure: set figure(numbering: n => chap-num(n))
  show figure: set block(breakable: false)        // caption never separated from its figure
  // A figure can never be taller than ~70% of the text area: tall diagrams
  // are scaled down (keeping proportions) instead of running off the page.
  show image: it => layout(size => {
    let natural = measure(it, width: size.width)
    let limit = 0.70 * (page.height - 2in)
    if natural.height > limit { image(it.source, height: limit) } else { it }
  })
  // Figures and tables float to the top/bottom of a page instead of leaving
  // a half-empty page behind them; text flows around to fill the space.
  // On a chapter's opening page a figure may only float to the bottom, so it
  // can never appear above the chapter title (same rule LaTeX uses).
  let float-fig(it) = context {
    let opener = query(heading.where(level: 1).before(here()))
    let chapter-page = if opener.len() > 0 { opener.last().location().page() } else { 0 }
    let where = if here().page() == chapter-page { bottom } else { auto }
    std.place(where, float: true, clearance: 1.5em, it)
  }
  show figure.where(kind: "quarto-float-fig"): float-fig
  show figure.where(kind: "quarto-float-tbl"): float-fig
  show figure.where(kind: "quarto-float-alg"): set align(left)
  show figure.where(kind: "quarto-float-alg"): it => block(
    width: 100%, stroke: (top: 0.8pt, bottom: 0.8pt), inset: (y: 6pt), it)
  show figure.caption: set text(size: size - 1pt)
  show figure.caption: set par(leading: 0.4em)
  show figure.caption: it => context [
    #text(weight: "bold")[#it.supplement #it.counter.display(it.numbering):] #it.body
  ]
  set table(stroke: 0.5pt, inset: 6pt)
  show table: set par(leading: 0.35em, justify: false)
  // Header cells are made bold by filters/tables.lua; centre all cells vertically.
  set table.cell(align: horizon)

  // ── Front lists and references look like chapters, unnumbered ──
  show outline: set heading(numbering: none)
  show outline: set par(leading: 0.5em)
  set outline.entry(fill: repeat[.#h(4pt)])
  show outline.entry.where(level: 1): it => {
    if it.element.func() == heading { v(0.6em, weak: true); strong(it) } else { it }
  }
  set bibliography(title: [References])   // style comes from report.yml citation_style

  // ── Cover page (no page number), from profile.cover ───────
  {
    set align(center)
    set par(justify: false, leading: 0.45em, spacing: 0.6em)
    v(0.5cm)
    for item in P.cover {
      if "text" in item {
        let body = md(item.text)
        if item.at("upper", default: false) { body = upper(body) }
        let t = text(size: item.at("size", default: 12) * 1pt, body)
        if item.at("bold", default: false) { t = strong(t) }
        if item.at("italic", default: false) { t = emph(t) }
        par(t)
      } else if "space" in item {
        v(item.space)
      } else if "fill" in item {
        v(1fr)
      } else if "logo" in item {
        image(item.path, width: item.logo)
      }
    }
  }
  pagebreak()

  // Front matter: page numbers (roman by default) start here.
  set page(numbering: num-style(P.page.numbers.front))
  counter(page).update(1)

  // ── Front pages (certificate, declaration, ...), from profile.pages ──
  for pg in P.pages {
    align(center, block(below: 1.2cm, text(size: P.headings.h1 * 1pt, weight: "bold", upper(pg.title))))
    if "image" in pg {
      align(center, image(pg.image, height: 80%, fit: "contain"))
    }
    if "text" in pg { md(pg.text) }
    if pg.sign.len() > 0 {
      v(1fr)
      if pg.sign.len() == 1 {
        align(right, sign-block(pg.sign.at(0)))
      } else {
        grid(columns: (1fr,) * pg.sign.len(), ..pg.sign.map(sign-block))
      }
    }
    if pg.place != none {
      v(1cm)
      [Place: #md(pg.place) #h(1fr) Date: #h(3cm)]
    }
    pagebreak()
  }

  doc
}
