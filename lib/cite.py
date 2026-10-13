"""Citations: fetch, verify, normalise and review references.

    report cite add "<title | DOI | arXiv id>" [...]   look up and append to refs.bib
    report cite verify [--fix]                         check every refs.bib entry against
                                                       Crossref / DBLP / OpenAlex / arXiv
    report cite review                                 write _build/citations.md: each citing
                                                       sentence next to the paper's abstract

Rules this module enforces:
  - References are fetched, never written from memory.
  - The published version wins over an arXiv preprint of the same paper.
  - Every verified entry carries `verified = {<source> <date>}`; entries with no
    match stay unverified until you check them by hand and set
    `verified = {manual <date>}`.
  - Abstracts are saved in report/refs/abstracts.json for the claim check.
"""
import difflib
import hashlib
import http.client
import json
import os
import re
import time
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

import bibtexparser
from bibtexparser.bparser import BibTexParser
from bibtexparser.bwriter import BibTexWriter
from bibtexparser.customization import convert_to_unicode

KIT = Path(__file__).resolve().parent.parent
CACHE = KIT / "cache/http"
UA = "report-kit/0.2 (citation verifier; personal academic use)"
MATCH = 0.92          # normalised-title similarity needed to call two titles the same paper
TODAY = date.today().isoformat()


# ── HTTP with cache and polite retries ───────────────────────
BLOCKED = set()   # hosts that answered with a bot challenge this session


def looks_valid(text, accept):
    head = text.lstrip()[:200].lower()
    if "html" not in accept and (head.startswith("<!doctype html") or head.startswith("<html")):
        return False
    if "json" in accept:
        try:
            json.loads(text)
        except json.JSONDecodeError:
            return False
    return True


def fetch(url, accept="application/json", headers=None):
    """GET with an on-disk cache. Only valid answers are cached; a host that
    answers with a bot-challenge page is skipped for the rest of the run."""
    CACHE.mkdir(parents=True, exist_ok=True)
    key = hashlib.sha1(f"{url}|{accept}".encode()).hexdigest()
    path = CACHE / key
    if path.exists():
        return path.read_text()
    host = urllib.parse.urlparse(url).netloc
    if host in BLOCKED:
        return None
    hdrs = {"User-Agent": UA, "Accept": accept, **(headers or {})}
    delay = 1.5
    for _ in range(6):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=hdrs), timeout=30) as r:
                text = r.read().decode("utf-8", "replace")
            if not looks_valid(text, accept):
                BLOCKED.add(host)
                print(f"  (skipping {host}: it answered with a bot check, not data)")
                return None
            path.write_text(text)
            return text
        except urllib.error.HTTPError as e:
            if e.code == 404:
                path.write_text("")
                return ""
            if e.code in (429, 500, 502, 503, 504):
                time.sleep(delay)
                delay *= 2
                continue
            return None
        except (urllib.error.URLError, TimeoutError, ConnectionError, OSError, http.client.HTTPException):
            time.sleep(delay)
            delay *= 2
    return None   # network trouble: not cached, try again next time


def fetch_json(url, **kw):
    text = fetch(url, **kw)
    try:
        return json.loads(text) if text else None
    except json.JSONDecodeError:
        return None


# ── text helpers ─────────────────────────────────────────────
LATEX_ACCENTS = {'"': "̈", "'": "́", "`": "̀", "^": "̂", "~": "̃",
                 "c": "̧", "v": "̌", "=": "̄", ".": "̇", "u": "̆"}


def delatex(s):
    """{\\"u} / \\"{u} / {\\'e} -> ü / é; drop protective braces."""
    def acc(m):
        return unicodedata.normalize("NFC", m.group(2) + LATEX_ACCENTS.get(m.group(1), ""))
    s = re.sub(r"\\([ij])(?![A-Za-z])", r"\1", s or "")          # dotless \i, \j
    s = re.sub(r"\{?\\([\"'`^~cv=.u])\{?\s*([A-Za-z])\}?\}?", acc, s)
    s = s.replace("\\&", "&").replace("--", "–")
    return re.sub(r"[{}]", "", s).strip()


def norm_title(t):
    t = unicodedata.normalize("NFKD", delatex(t)).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", " ", t).strip()


def title_sim(a, b):
    return difflib.SequenceMatcher(None, norm_title(a), norm_title(b)).ratio()


PARTICLES = {"van", "von", "de", "da", "di", "del", "der", "du", "le", "la", "den", "ten", "ter", "dos", "das"}


def split_name(full):
    """'Blaise Agüera y Arcas' -> ('Blaise', 'Agüera y Arcas'); 'Ludwig van Beethoven' ->
    ('Ludwig', 'van Beethoven'). DBLP disambiguation suffixes ('Li Zhang 0001') are dropped."""
    toks = re.sub(r"\s+\d{4}$", "", full.strip()).split()
    if len(toks) == 1:
        return "", toks[0]
    for i, t in enumerate(toks[1:], 1):
        if t == "y" and i >= 2:
            return " ".join(toks[: i - 1]), " ".join(toks[i - 1:])
        if t.lower() in PARTICLES and i < len(toks) - 1:
            return " ".join(toks[:i]), " ".join(toks[i:])
    return " ".join(toks[:-1]), toks[-1]


def bib_author(given, family):
    fam = f"{{{family}}}" if " " in family else family
    return f"{fam}, {given}" if given else fam


def parse_bib_authors(field):
    """BibTeX author field -> [(given, family)]."""
    out = []
    for a in re.split(r"\s+and\s+", field or ""):
        a = delatex(a).strip()
        if not a or a.lower() == "others":
            continue
        if "," in a:
            fam, given = a.split(",", 1)
            out.append((given.strip(), fam.strip()))
        else:
            out.append(split_name(a))
    return out


def normalize_author_field(field):
    """'Arcas, Blaise Agüera y' -> '{Agüera y Arcas}, Blaise'; 'Beethoven, Ludwig van' ->
    '{van Beethoven}, Ludwig'. Particles stranded in the given name move to the surname."""
    out = []
    for given, family in parse_bib_authors(field):
        toks = given.split()
        if len(toks) >= 2 and toks[-1] == "y":
            given, family = " ".join(toks[:-2]), f"{toks[-2]} y {family}"
        elif len(toks) >= 2 and toks[-1].lower() in PARTICLES:
            given, family = " ".join(toks[:-1]), f"{toks[-1]} {family}"
        out.append(bib_author(given, family))
    if re.search(r"\band\s+others\s*$", field or ""):
        out.append("others")
    return " and ".join(out)


SMALL_KEEP = re.compile(r"^(\{.*\}|[A-Z]{2,}.*|.*\d.*)$")


def sentence_case(title):
    """APA/most author-date styles: capitalise the first word and the first word after a
    colon; lower-case the rest unless braced ({FedAvg}) or an acronym/number."""
    parts = re.split(r"(\{[^{}]*\}|\s+)", title)
    out, first = [], True
    for part in parts:
        if not part or part.isspace():
            out.append(part or "")
            continue
        if part.startswith("{"):
            out.append(part)
            first = part.rstrip("}").endswith(":")
            continue
        if first:   # "Data-Driven" -> "Data-driven"
            head, *rest = part.split("-")
            part = "-".join([head] + [r if SMALL_KEEP.match(r) else r.lower() for r in rest])
            out.append(part)
        else:
            out.append(part if SMALL_KEEP.match(part) else part.lower())
        first = part.endswith(":")
    return "".join(out)


def protect_title(t):
    """Brace words whose capitals must survive citation styles: FedAvg, IoT, 5G, BERT, Non-IID."""
    t = delatex(t)

    def needs_guard(part):   # IID, IoT, FedAvg, 5G, iPhone -- not Privacy, Data, Driven, 2020
        return bool(re.search(r"^.+[A-Z]|\d", part)) and not part.isdigit()

    def guard(m):
        w = m.group(0)
        return f"{{{w}}}" if any(needs_guard(p) for p in w.split("-")) else w
    return re.sub(r"[A-Za-z0-9][A-Za-z0-9\-]*", guard, t)


# ── candidates from each source ──────────────────────────────
# A candidate is a dict: source, id, title, authors [(given, family)], year, type,
# venue, volume, number, pages, doi, arxiv, publisher, preprint (bool), abstract

def crossref_item(m):
    typ = m.get("type", "")
    venue = (m.get("container-title") or [""])[0]
    years = {m[k]["date-parts"][0][0] for k in ("published-print", "published-online", "issued")
             if (m.get(k, {}).get("date-parts") or [[None]])[0][0]}
    year = next((m[k]["date-parts"][0][0] for k in ("published-print", "issued", "published-online")
                 if (m.get(k, {}).get("date-parts") or [[None]])[0][0]), None)
    ab = re.sub(r"<[^>]+>", " ", m.get("abstract", "") or "")
    return {
        "source": "crossref", "id": m.get("DOI"), "title": (m.get("title") or [""])[0],
        "authors": [(a.get("given", ""), a.get("family", a.get("name", ""))) for a in m.get("author", [])],
        "year": year, "years": sorted(years), "venue": venue, "volume": m.get("volume"), "number": m.get("issue"),
        "pages": (m.get("page") or "").replace("-", "--") or None, "doi": m.get("DOI"),
        "publisher": m.get("publisher"),
        "type": {"journal-article": "article", "proceedings-article": "inproceedings",
                 "book-chapter": "incollection", "book": "book"}.get(typ, "misc"),
        "preprint": typ == "posted-content", "abstract": re.sub(r"\s+", " ", ab).strip() or None,
    }


def crossref_doi(doi):
    d = fetch_json(f"https://api.crossref.org/works/{urllib.parse.quote(doi)}")
    return [crossref_item(d["message"])] if d and d.get("message") else []


def crossref_search(title):
    q = urllib.parse.urlencode({"query.bibliographic": title, "rows": 5})
    d = fetch_json(f"https://api.crossref.org/works?{q}")
    return [crossref_item(m) for m in (d or {}).get("message", {}).get("items", [])]


def dblp_search(title):
    q = urllib.parse.urlencode({"q": title, "format": "json", "h": 10})
    d = fetch_json(f"https://dblp.org/search/publ/api?{q}")
    hits = (((d or {}).get("result") or {}).get("hits") or {}).get("hit") or []
    out = []
    for h in hits:
        info = h["info"]
        authors = info.get("authors", {}).get("author", [])
        authors = authors if isinstance(authors, list) else [authors]
        preprint = info.get("type", "").startswith("Informal") or info.get("venue") == "CoRR"
        out.append({
            "source": "dblp", "id": info.get("key"), "title": delatex(info.get("title", "")).rstrip("."),
            "authors": [split_name(a["text"]) for a in authors], "year": int(info["year"]) if info.get("year") else None,
            "venue": info.get("venue"), "doi": info.get("doi"), "preprint": preprint,
            "type": "article" if info.get("type") == "Journal Articles" else
                    "inproceedings" if info.get("type", "").startswith("Conference") else "misc",
            "arxiv": (re.search(r"abs/(\d{4}\.\d{4,5})", info.get("ee", "")) or [None, None])[1] if preprint else None,
        })
    return out


def dblp_bib(key):
    """Full DBLP BibTeX record (full booktitle, pages, publisher) for a DBLP key."""
    text = fetch(f"https://dblp.org/rec/{key}.bib?param=1", accept="text/plain")
    if not text:
        return None
    db = bibtexparser.loads(text, BibTexParser(common_strings=True))
    return db.entries[0] if db.entries else None


def openalex(title=None, doi=None):
    sel = "id,doi,title,publication_year,primary_location,locations,authorships,abstract_inverted_index,type"
    if doi:
        d = fetch_json(f"https://api.openalex.org/works/doi:{urllib.parse.quote(doi)}?select={sel}")
        works = [d] if d and d.get("id") else []
    else:
        q = urllib.parse.urlencode({"search": title, "per-page": 5, "select": sel})
        works = ((fetch_json(f"https://api.openalex.org/works?{q}") or {}).get("results")) or []
    out = []
    for w in works:
        inv = w.get("abstract_inverted_index") or {}
        words = sorted((p, t) for t, ps in inv.items() for p in ps)
        src = ((w.get("primary_location") or {}).get("source") or {})
        out.append({
            "source": "openalex", "id": w.get("id"), "title": w.get("title") or "",
            "authors": [split_name(a["author"]["display_name"]) for a in w.get("authorships", [])],
            "year": w.get("publication_year"), "venue": src.get("display_name"),
            "doi": (w.get("doi") or "").replace("https://doi.org/", "") or None,
            "preprint": w.get("type") == "preprint", "type": "misc",
            "abstract": " ".join(t for _, t in words) or None,
            "venues": [((l.get("source") or {}).get("display_name") or "", (l.get("source") or {}).get("type") or "")
                       for l in w.get("locations", [])],
        })
    return out


def arxiv(arxiv_id):
    text = fetch(f"https://export.arxiv.org/api/query?id_list={arxiv_id}", accept="application/atom+xml")
    if not text:
        return []
    ns = {"a": "http://www.w3.org/2005/Atom"}
    e = ET.fromstring(text).find("a:entry", ns)
    if e is None or e.find("a:title", ns) is None:
        return []
    return [{
        "source": "arxiv", "id": arxiv_id, "title": re.sub(r"\s+", " ", e.find("a:title", ns).text).strip(),
        "authors": [split_name(a.find("a:name", ns).text) for a in e.findall("a:author", ns)],
        "year": int(e.find("a:published", ns).text[:4]), "venue": "arXiv", "preprint": True, "type": "misc",
        "arxiv": arxiv_id, "abstract": re.sub(r"\s+", " ", e.find("a:summary", ns).text).strip(),
    }]


def semantic_scholar_abstract(title):
    key = os.environ.get("SEMANTIC_SCHOLAR_API_KEY")
    if not key:
        return None
    q = urllib.parse.urlencode({"query": title, "fields": "title,abstract", "limit": 3})
    d = fetch_json(f"https://api.semanticscholar.org/graph/v1/paper/search?{q}", headers={"x-api-key": key})
    for p in (d or {}).get("data", []):
        if p.get("abstract") and title_sim(p["title"], title) >= MATCH:
            return p["abstract"]
    return None


# ── resolve: best published match for a query ───────────────
DOI_RE = re.compile(r"\b(10\.\d{4,9}/[^\s\"<>]+)")
ARXIV_RE = re.compile(r"(?:arxiv[:/ ]*)?(\d{4}\.\d{4,5})(v\d+)?", re.I)


def resolve(query=None, title=None, doi=None):
    """Return (best_candidate, all_candidates). Published beats preprint; DOI lookups are exact."""
    if query:
        m = DOI_RE.search(query)
        a = ARXIV_RE.fullmatch(query.strip()) or (re.search(r"arxiv", query, re.I) and ARXIV_RE.search(query))
        if m:
            doi = m.group(1).rstrip(".")
        elif a:
            cands = arxiv(a.group(1))
            if not cands:
                return None, []
            title = cands[0]["title"]
            best, more = resolve(title=title)
            if best and not best["preprint"]:
                return best, cands + more
            return cands[0], cands + more
        else:
            title = query
    if doi and not title:
        cands = crossref_doi(doi)
        if "arxiv" in doi.lower() and cands:      # arXiv DOI: look for the published version
            best, more = resolve(title=cands[0]["title"])
            if best and not best["preprint"]:
                return best, cands + more
        return (cands[0] if cands else None), cands
    cands = dblp_search(title) + crossref_search(title)
    matching = [c for c in cands if c["title"] and title_sim(c["title"], title) >= MATCH]
    if not matching:
        oa = openalex(title=title)
        matching = [c for c in oa if title_sim(c["title"], title) >= MATCH]
        cands += oa
    if not matching:
        return None, cands
    rank = {"dblp": 0, "crossref": 1, "openalex": 2, "arxiv": 3}
    matching.sort(key=lambda c: (c["preprint"], rank[c["source"]]))
    best = matching[0]
    if best["source"] != "crossref" and best.get("doi") and "arxiv" not in best["doi"].lower():
        full = crossref_doi(best["doi"])
        if full and title_sim(full[0]["title"], best["title"]) >= MATCH:
            full[0]["abstract"] = full[0].get("abstract") or best.get("abstract")
            best = full[0]
    return best, cands


def find_abstract(cand, title):
    if cand and cand.get("abstract"):
        return cand["abstract"], cand["source"]
    if cand and cand.get("doi"):
        for c in openalex(doi=cand["doi"]):
            if c.get("abstract"):
                return c["abstract"], "openalex"
    for c in openalex(title=title):
        if c.get("abstract") and title_sim(c["title"], title) >= MATCH:
            return c["abstract"], "openalex"
    for c in dblp_search(title):
        if c.get("arxiv") and title_sim(c["title"], title) >= MATCH:
            for a in arxiv(c["arxiv"]):
                return a["abstract"], "arxiv"
    ab = semantic_scholar_abstract(title)
    return (ab, "semanticscholar") if ab else (None, None)


# ── candidate -> BibTeX entry ────────────────────────────────
def to_entry(cand, key=None):
    if cand["source"] == "dblp" and cand.get("id"):
        full = dblp_bib(cand["id"])
        if full:
            e = {k: v for k, v in full.items() if k in (
                "title", "author", "booktitle", "journal", "volume", "number", "pages", "year",
                "publisher", "series", "doi", "eprint", "eprinttype")}
            e["ENTRYTYPE"] = full["ENTRYTYPE"]
            e["author"] = " and ".join(bib_author(g, f) for g, f in parse_bib_authors(full["author"]))
            if e.get("booktitle"):    # DBLP appends place and date; keep the name
                e["booktitle"] = re.split(r",\s*(?:\d{1,2}[-–]\d{1,2}|[A-Z][a-z]+ \d{1,2}[-–])", delatex(e["booktitle"]))[0]
            for k in ("journal", "series", "publisher"):
                if e.get(k):
                    e[k] = delatex(e[k])
        else:
            e = None
    else:
        e = None
    if e is None:
        e = {"ENTRYTYPE": cand.get("type") or "misc",
             "author": " and ".join(bib_author(g, f) for g, f in cand["authors"])}
        e["title"] = cand["title"]
        if cand.get("year"):
            e["year"] = str(cand["year"])
        venue_field = {"article": "journal", "inproceedings": "booktitle", "incollection": "booktitle"}.get(e["ENTRYTYPE"])
        if venue_field and cand.get("venue"):
            e[venue_field] = cand["venue"]
        for f in ("volume", "number", "pages", "doi", "publisher"):
            if cand.get(f):
                e[f] = str(cand[f])
        if cand.get("arxiv"):
            e.update(eprint=cand["arxiv"], eprinttype="arXiv", note=f"arXiv:{cand['arxiv']}")
    e["title"] = protect_title(e["title"])
    e["author"] = normalize_author_field(e["author"])
    e["verified"] = f"{cand['source']} {TODAY}"
    if not key:
        fam = norm_title(cand["authors"][0][1]).replace(" ", "") if cand["authors"] else "anon"
        first = next((w for w in norm_title(cand["title"]).split() if w not in {"a", "an", "the", "on", "of"}), "x")
        key = f"{fam}{cand.get('year') or ''}{first}"
    e["ID"] = key
    return e


# ── refs.bib I/O ─────────────────────────────────────────────
def load_bib(rdir):
    path = rdir / "refs.bib"
    parser = BibTexParser(common_strings=True, ignore_nonstandard_types=False)
    parser.homogenize_fields = True
    return bibtexparser.loads(path.read_text() if path.exists() else "", parser)


def save_bib(rdir, db):
    w = BibTexWriter()
    w.indent = "  "
    w.order_entries_by = None
    w.display_order = ["title", "author", "booktitle", "journal", "volume", "number", "pages", "year",
                       "publisher", "doi", "eprint", "eprinttype", "note", "verified"]
    (rdir / "refs.bib").write_text(bibtexparser.dumps(db, w))


def load_abstracts(rdir):
    p = rdir / "refs/abstracts.json"
    return json.loads(p.read_text()) if p.exists() else {}


def save_abstracts(rdir, abstracts):
    p = rdir / "refs/abstracts.json"
    p.parent.mkdir(exist_ok=True)
    p.write_text(json.dumps(abstracts, indent=1, ensure_ascii=False, sort_keys=True))


def write_typst_bib(rdir, style):
    """Give Typst the titles the way Pandoc's citeproc sees them: sentence case,
    braced words kept (Pandoc lowercases BibTeX titles before applying a CSL style;
    Typst doesn't, so without this the PDF and DOCX reference lists would differ).
    Returns the path to use, relative to rdir."""
    if not (rdir / "refs.bib").exists():
        return "refs.bib"
    db = load_bib(rdir)
    for e in db.entries:
        if e.get("title"):
            e["title"] = sentence_case(e["title"])
    (rdir / "refs").mkdir(exist_ok=True)
    (rdir / "refs/typst.bib").write_text(bibtexparser.dumps(db, BibTexWriter()))
    return "refs/typst.bib"


# ── commands ─────────────────────────────────────────────────
def cmd_add(rdir, queries):
    db = load_bib(rdir)
    abstracts = load_abstracts(rdir)
    have = {e["ID"]: e for e in db.entries}
    for q in queries:
        best, _ = resolve(query=q)
        if not best:
            print(f"  ✗ not found: {q!r}. Add it by hand with verified = {{manual {TODAY}}} once you've checked it.")
            continue
        dup = next((k for k, e in have.items() if title_sim(e.get("title", ""), best["title"]) >= MATCH), None)
        if dup:
            print(f"  · already in refs.bib as [@{dup}]: {best['title']}")
            continue
        e = to_entry(best)
        while e["ID"] in have:
            e["ID"] += "b"
        db.entries.append(e)
        have[e["ID"]] = e
        ab, src = find_abstract(best, best["title"])
        if ab:
            abstracts[e["ID"]] = {"source": src, "abstract": ab}
        print(f"  ✓ [@{e['ID']}] {delatex(e['title'])} ({e.get('year')}, via {best['source']}"
              f"{', preprint' if best['preprint'] else ''}){'' if ab else ' — no abstract found'}")
        if best["preprint"]:
            pubs = {v for c in openalex(title=best["title"]) if title_sim(c["title"], best["title"]) >= MATCH
                    for v, t in (c.get("venues") or []) if t in ("conference", "journal")}
            if pubs:
                print(f"      ! a published version exists ({'; '.join(sorted(pubs))}). Cite that instead: "
                      f"set the entry type, booktitle/journal and year by hand and add verified = {{manual {TODAY}}}")
    save_bib(rdir, db)
    save_abstracts(rdir, abstracts)


def same_person(family, cand_author):
    """Surname match that survives Chinese name order and initials-only records."""
    fam = norm_title(family)
    full = norm_title(" ".join(cand_author))
    return fam in full.split() or fam.replace(" ", "") in full.replace(" ", "")


def compare(entry, cand):
    """Hard differences between a refs.bib entry and the authoritative record,
    plus soft notes that don't block verification."""
    hard, soft = [], []
    if title_sim(entry.get("title", ""), cand["title"]) < MATCH:
        hard.append(f"title: “{delatex(entry.get('title', ''))}” vs “{cand['title']}”")
    y = entry.get("year")
    if cand.get("year") and y and not cand["preprint"]:
        if cand.get("years"):        # Crossref: accept the print or the online year, nothing else
            if int(y) not in cand["years"]:
                hard.append(f"year: {y} vs {'/'.join(map(str, cand['years']))}")
            elif len(cand["years"]) > 1:
                soft.append(f"published online {min(cand['years'])}, in print {max(cand['years'])}")
        elif str(cand["year"]) != str(y):
            (soft if abs(int(cand["year"]) - int(y)) == 1 else hard).append(f"year: {y} vs {cand['year']}")
    mine = parse_bib_authors(entry.get("author", ""))
    if cand["authors"] and mine and not same_person(mine[0][1], cand["authors"][0]):
        hard.append(f"first author: {mine[0][1]} vs {' '.join(cand['authors'][0])}")
    if entry.get("doi") and cand.get("doi") and entry["doi"].lower() != cand["doi"].lower():
        hard.append(f"doi: {entry['doi']} vs {cand['doi']}")
    return hard, soft


# Venues without DOIs whose yearly proceedings index is public: a title found on
# the index for the entry's year confirms both venue and year.
PROCEEDINGS = [
    (r"machine learning and systems|\bmlsys\b|\bsysml\b", "https://proceedings.mlsys.org/paper_files/paper/{year}", "mlsys-proceedings"),
    (r"neural information processing systems|\bneurips\b|\bnips\b", "https://papers.nips.cc/paper_files/paper/{year}", "neurips-proceedings"),
]


def in_proceedings(entry):
    """Confirm a no-DOI conference paper against the venue's official yearly index."""
    venue = delatex(entry.get("booktitle") or entry.get("journal") or "")
    year = entry.get("year")
    for pat, url, label in PROCEEDINGS:
        if year and re.search(pat, venue, re.I):
            page = fetch(url.format(year=year), accept="text/html")
            if not page:
                return None
            titles = re.findall(r'title="paper title"[^>]*>([^<]+)</a>', page)
            if any(title_sim(t, entry.get("title", "")) >= MATCH for t in titles):
                return label
    return None


def venue_listed(entry, cand):
    """Does an OpenAlex record list the entry's own venue among its locations?"""
    mine = words(delatex(entry.get("booktitle") or entry.get("journal") or ""))
    mine -= {"proceeding", "international", "conference", "annual", "journal", "ieee", "acm"}
    if not mine:
        return False
    for v, _ in cand.get("venues") or []:
        theirs = words(v)
        if theirs and len(mine & theirs) / len(mine) >= 0.5:
            return True
    return False


def cmd_verify(rdir, fix=False):
    db = load_bib(rdir)
    abstracts = load_abstracts(rdir)
    counts = {"ok": 0, "fixed": 0, "mismatch": 0, "unconfirmed": 0, "unverified": 0, "manual": 0}
    new_entries = []
    for e in db.entries:
        if str(e.get("verified", "")).startswith("manual"):
            counts["manual"] += 1
            new_entries.append(e)
            print(f"  · [@{e['ID']}] checked by hand ({e['verified']})")
            continue
        title = delatex(e.get("title", ""))
        best, cands = resolve(doi=e.get("doi")) if e.get("doi") else resolve(query=e.get("eprint") or title)
        if best and e.get("doi") and title_sim(best["title"], title) < MATCH:
            best = None
        if best and not e.get("doi"):
            # Several records can share a title (journal issue vs book edition, preprint vs
            # paper): verify against the one that agrees with the entry, if any does.
            same = [c for c in cands if c["title"] and title_sim(c["title"], title) >= MATCH and not c["preprint"]]
            agreeing = [c for c in same if not compare(e, c)[0]]
            if agreeing and compare(e, best)[0]:
                best = agreeing[0]
                if best["source"] != "crossref" and best.get("doi"):
                    best = (crossref_doi(best["doi"]) or [best])[0]
        if not best:
            counts["unverified"] += 1
            e.pop("verified", None)
            new_entries.append(e)
            print(f"  ✗ [@{e['ID']}] NOT FOUND in Crossref/DBLP/OpenAlex/arXiv: {title[:70]}\n"
                  f"      check it by hand, then add  verified = {{manual {TODAY}}}")
            continue
        hard, soft = compare(e, best)
        claims_published = e.get("ENTRYTYPE") in ("inproceedings", "article", "incollection") and \
            not re.search(r"arxiv", " ".join(str(v) for v in e.values()), re.I)
        if not hard and best["preprint"] and claims_published:
            # The paper exists, but only a preprint record was found.
            oa = [c for c in openalex(title=title) if title_sim(c["title"], title) >= MATCH]
            proc = in_proceedings(e)
            if proc:
                e["verified"] = f"{proc} {TODAY}"
                counts["ok"] += 1
                print(f"  ✓ [@{e['ID']}] {title[:60]} (found in the {proc.replace('-', ' ')} for {e.get('year')})")
            elif any(venue_listed(e, c) for c in oa):
                e["verified"] = f"openalex-venue {TODAY}"
                counts["ok"] += 1
                print(f"  ✓ [@{e['ID']}] {title[:60]} (venue listed in openalex)")
            else:
                e["verified"] = f"unconfirmed-venue {TODAY}"
                counts["unconfirmed"] += 1
                print(f"  ? [@{e['ID']}] paper exists ({best['source']}, preprint {best.get('year')}), but "
                      f"“{delatex(e.get('booktitle') or e.get('journal') or '')[:50]}” {e.get('year')} could not be "
                      f"confirmed automatically.\n      Check it once, then set  verified = {{manual {TODAY}}}")
            e["title"] = protect_title(e.get("title", ""))
            e["author"] = normalize_author_field(e.get("author", ""))
            new_entries.append(e)
        elif hard:
            if fix:
                fixed = to_entry(best, key=e["ID"])
                new_entries.append(fixed)
                counts["fixed"] += 1
                print(f"  ✎ [@{e['ID']}] fixed from {best['source']}: {'; '.join(hard)}")
            else:
                counts["mismatch"] += 1
                e.pop("verified", None)
                new_entries.append(e)
                print(f"  ! [@{e['ID']}] differs from {best['source']}: {'; '.join(hard)}")
        else:
            e["verified"] = f"{best['source']} {TODAY}"
            e["title"] = protect_title(e.get("title", ""))
            e["author"] = normalize_author_field(e.get("author", ""))
            if fix and best["source"] in ("crossref", "dblp") and not best["preprint"]:
                for f in ("doi", "volume", "number", "pages"):          # fill gaps, never overwrite
                    if best.get(f) and not e.get(f):
                        e[f] = str(best[f])
            new_entries.append(e)
            counts["ok"] += 1
            note = f" — note: {'; '.join(soft)}" if soft else ""
            print(f"  ✓ [@{e['ID']}] {title[:60]} ({best['source']}){note}")
        if e["ID"] not in abstracts:
            ab, src = find_abstract(best, best["title"])
            if ab:
                abstracts[e["ID"]] = {"source": src, "abstract": ab}
    # Duplicates (e.g. an arXiv preprint and its published version both cited).
    for i, a in enumerate(new_entries):
        for b in new_entries[i + 1:]:
            if title_sim(a.get("title", ""), b.get("title", "")) >= MATCH:
                print(f"  ! [@{a['ID']}] and [@{b['ID']}] are the same paper: keep one")
    db.entries = new_entries
    save_bib(rdir, db)
    save_abstracts(rdir, abstracts)
    print(f"\n  {counts['ok']} verified, {counts['fixed']} fixed, {counts['mismatch']} differ (run with --fix), "
          f"{counts['unconfirmed']} venue unconfirmed, {counts['unverified']} not found, "
          f"{counts['manual']} checked by hand")
    missing = [e["ID"] for e in new_entries if e["ID"] not in abstracts]
    if missing:
        print(f"  no abstract for: {', '.join(missing)} (claim check skipped for these)")


# ── claim support ────────────────────────────────────────────
STOP = set("""a an the and or of in on for to with by from as at is are was were be been this that these
those it its their our we they which who whom whose than then so such can could may might will would
not no also only into over under more most less least very using use used based via while when where
how what each other between both all any some per one two three new show shows shown""".split())


def words(text):
    return {w[:-1] if w.endswith("s") and len(w) > 4 else w
            for w in re.findall(r"[a-z][a-z0-9\-]{2,}", text.lower()) if w not in STOP}


def citation_contexts(rdir):
    """[(key, file, sentence)] for every [@key] in front/ and chapters/."""
    out = []
    for f in sorted((rdir / "front").glob("*.md")) + sorted((rdir / "chapters").glob("*.md")):
        text = re.sub(r"```.*?```", " ", f.read_text(), flags=re.S)
        for para in re.split(r"\n\s*\n", text):
            para = " ".join(para.split())
            for sent in re.split(r"(?<=[.!?])\s+(?=[A-Z])", para):
                for group in re.findall(r"\[([^\]]*@[^\]]*)\]", sent):
                    for key in re.findall(r"@([\w:-]+)", group):
                        out.append((key, f.relative_to(rdir).as_posix(), sent))
    return out


def expand_abbreviations(text, glossary):
    for abbr, full in (glossary or {}).items():
        text = re.sub(rf"\b{re.escape(abbr)}\b", f"{full} {abbr}", text)
    return text


def support_score(sentence, entry, abstract, glossary=None):
    """Share of the claim's content words found in the cited paper's title + abstract.
    A low score is a hint to look, not proof of a bad citation."""
    claim = words(expand_abbreviations(re.sub(r"\[[^\]]*@[^\]]*\]", " ", sentence), glossary))
    paper = words(delatex(entry.get("title", "")) + " " + (abstract or ""))
    shared = claim & paper
    return (len(shared) / max(len(claim), 1)), sorted(shared)


def cmd_review(rdir):
    import check
    glossary = check.read_glossary(rdir)
    db = {e["ID"]: e for e in load_bib(rdir).entries}
    abstracts = load_abstracts(rdir)
    lines = ["# Citation review", "",
             "Each citing sentence next to the abstract of the paper it cites. Check that the paper",
             "really says what the sentence claims. Low overlap is flagged ⚠ — it is a hint, not proof.", ""]
    for key in sorted({k for k, _, _ in citation_contexts(rdir)}):
        e = db.get(key)
        ab = (abstracts.get(key) or {}).get("abstract")
        lines.append(f"## [@{key}] {delatex(e.get('title', '')) if e else '(not in refs.bib)'}")
        lines.append(f"*{e.get('verified', 'NOT VERIFIED') if e else ''}*\n")
        lines.append(f"> {ab}\n" if ab else "> (no abstract available)\n")
        for k, f, sent in citation_contexts(rdir):
            if k != key:
                continue
            score, shared = support_score(sent, e or {}, ab, glossary)
            flag = "⚠ " if ab and score < 0.15 else ""
            lines.append(f"- {flag}`{f}` ({score:.0%} overlap: {', '.join(shared[:8]) or '—'}): {sent}")
        lines.append("")
    out = rdir / "_build/citations.md"
    out.parent.mkdir(exist_ok=True)
    out.write_text("\n".join(lines))
    print(f"  wrote {out.relative_to(rdir.parent)}")


# ── checks used by every build ───────────────────────────────
def checks(rdir, final):
    gap = "FAIL" if final else "WARN"
    out = []
    if not (rdir / "refs.bib").exists():
        return out
    entries = {e["ID"]: e for e in load_bib(rdir).entries}
    abstracts = load_abstracts(rdir)
    contexts = citation_contexts(rdir)
    cited = {k for k, _, _ in contexts}
    for key in sorted(cited - set(entries)):
        if not re.match(r"(fig|tbl|eq|alg|sec)-", key):
            out.append(("FAIL", f"[@{key}] is cited but not in refs.bib"))
    for key in sorted(cited & set(entries)):
        v = str(entries[key].get("verified", ""))
        if not v:
            out.append((gap, f"[@{key}] is not verified: run `report cite verify`"))
        elif v.startswith("unconfirmed"):
            out.append((gap, f"[@{key}] exists, but its venue/year is unconfirmed: check once, "
                             f"then set verified = {{manual {TODAY}}} in refs.bib"))
    import check
    glossary = check.read_glossary(rdir)
    weak = 0
    for key, f, sent in contexts:
        ab = (abstracts.get(key) or {}).get("abstract")
        if key in entries and ab:
            score, _ = support_score(sent, entries[key], ab, glossary)
            if score < 0.15:
                weak += 1
                out.append(("WARN", f"[@{key}] in {f} may not support: “{sent[:90]}…”"))
    if weak:
        out.append(("INFO", "side-by-side view of citations and abstracts: report cite review"))
    uncited = sorted(set(entries) - cited)
    if uncited:
        out.append(("INFO", f"in refs.bib but never cited (left out of the list): {', '.join(uncited)}"))
    return out
