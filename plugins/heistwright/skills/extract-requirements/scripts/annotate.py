# /// script
# requires-python = ">=3.10"
# dependencies = ["pymupdf>=1.24", "python-docx>=1.2", "openpyxl>=3.1"]
# ///
"""Annotate a copy of the source document with the sources cited in the requirements Excel file.

    uv run annotate.py <source.pdf|source.docx> <requirements.xlsx>

Writes `<source name> - Annotated.<pdf|docx>` next to the Excel file, where every cited
passage is highlighted and commented with its requirement ID, and links each Location cell
in the Excel file to the requirement's first source. Quotes that cannot be found are
reported, never guessed. Safe to run again: it always starts from the source document.
"""
from __future__ import annotations

import re
import sys
from copy import copy, deepcopy
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.worksheet.hyperlink import Hyperlink

AUTHOR = "extract-requirements"
LANG_WORDS = {
    "en": {"also": "also", "qualifier": "qualifier", "root": "Requirements"},
    "no": {"also": "også", "qualifier": "presisering", "root": "Krav"},
}
HEADERS = {
    "id": {"id"},
    "text": {"original text", "opprinnelig tekst"},
    "loc": {"location", "plassering"},
}
PAGE_RE = re.compile(r"\b(?:page|side)\s+(\d+)(?:\s*[-–]\s*(\d+))?", re.I)
TABLE_MAX_GAP = 3  # chars allowed between cell values of a table row (e.g. a row number)
ELLIPSIS_MAX_GAP = 3000  # chars allowed for an omission marked with "..."
ANCHOR_LEN = 30  # start/end anchor length for long quotes broken by page breaks or bullets


# ---------------------------------------------------------------- text matching

_TRANS = str.maketrans({
    "‘": "'", "’": "'", "‚": "'", "“": '"', "”": '"', "„": '"',
    "«": '"', "»": '"', "–": "-", "—": "-", "−": "-", "‑": "-",
    "­": "", "•": "", "": "", "▪": "", "●": "", "…": "...",
})


def norm(c: str) -> str:
    """Normalised, whitespace-free form of one character (may be empty or several chars)."""
    return "".join(x for x in c.translate(_TRANS).lower() if not x.isspace())


def compact(s: str) -> str:
    return "".join(norm(c) for c in s)


@dataclass
class Pattern:
    segments: list[str]
    max_gap: int
    kind: str  # "text", "table", "ellipsis"


def pattern_for(quote: str) -> Pattern:
    if "|" in quote:
        segs = [compact(c) for c in quote.split("|")]
        return Pattern([s for s in segs if s], TABLE_MAX_GAP, "table")
    parts = re.split(r"\.\.\.|…|\[\.\.\.\]", quote)
    if len(parts) > 1:
        return Pattern([s for s in (compact(p) for p in parts) if s], ELLIPSIS_MAX_GAP, "ellipsis")
    return Pattern([compact(quote)], 0, "text")


def find_spans(text: str, pat: Pattern, limit: int = 5) -> list[tuple[int, int]]:
    """All (start, end) spans in `text` where the segments occur in order within max_gap."""
    if not pat.segments:
        return []
    spans, pos = [], 0
    while len(spans) < limit:
        s = text.find(pat.segments[0], pos)
        if s < 0:
            break
        end, ok = s + len(pat.segments[0]), True
        for seg in pat.segments[1:]:
            j = text.find(seg, end)
            if j < 0 or j - end > pat.max_gap:
                ok = False
                break
            end = j + len(seg)
        if ok:
            spans.append((s, end))
        pos = s + 1
    return spans


def find_anchored(text: str, quote_c: str) -> list[tuple[int, int]]:
    """Fallback for long quotes: match the first and last ANCHOR_LEN chars, close together."""
    if len(quote_c) < 3 * ANCHOR_LEN:
        return []
    head, tail = quote_c[:ANCHOR_LEN], quote_c[-ANCHOR_LEN:]
    spans, pos = [], 0
    while True:
        s = text.find(head, pos)
        if s < 0:
            return spans
        t = text.find(tail, s + ANCHOR_LEN)
        if 0 <= t and t + ANCHOR_LEN - s <= len(quote_c) * 1.5 + 200:
            spans.append((s, t + ANCHOR_LEN))
        pos = s + 1


# ---------------------------------------------------------------- Excel input

@dataclass
class Source:
    req: str
    k: int  # 1-based position among the requirement's sources
    n: int
    marker: str  # "*", "-" or "" (single source)
    quote: str
    location: str
    status: str = "pending"  # found | approximate | ambiguous | not found | external | too short
    note: str = ""
    anchor: object = None  # format-specific position, set when found

    @property
    def found(self) -> bool:
        return self.status in ("found", "approximate", "ambiguous")

    def pages(self) -> list[int]:
        m = PAGE_RE.search(self.location)
        if not m:
            return []
        a = int(m.group(1))
        b = int(m.group(2) or a)
        return list(range(a - 1, b))


@dataclass
class Requirement:
    rid: str
    row: int
    sources: list[Source] = field(default_factory=list)


def strip_quote(s: str) -> str:
    s = s.strip()
    if len(s) >= 2 and s[0] in "\"“«„'" and s[-1] in "\"”»“'":
        s = s[1:-1].strip()
    return s


def read_requirements(ws) -> tuple[dict[str, int], str, list[Requirement]]:
    header = {str(c.value or "").strip().lower(): c.column for c in ws[1]}
    cols = {}
    for key, names in HEADERS.items():
        match = next((col for name, col in header.items() if name in names), None)
        if match is None:
            sys.exit(f"Column not found in Excel header: one of {sorted(names)}")
        cols[key] = match
    loc_name = next(name for name, col in header.items() if col == cols["loc"])
    lang = "no" if loc_name == "plassering" else "en"
    reqs = []
    for row in range(2, ws.max_row + 1):
        rid = ws.cell(row, cols["id"]).value
        if not rid:
            continue
        quotes = [l for l in str(ws.cell(row, cols["text"]).value or "").splitlines() if l.strip()]
        locs = [l for l in str(ws.cell(row, cols["loc"]).value or "").splitlines() if l.strip()]
        req = Requirement(str(rid).strip(), row)
        for i, line in enumerate(quotes):
            line = line.strip()
            marker = line[0] if line[:2] in ("* ", "- ") else ""
            quote = strip_quote(line[2:] if marker else line)
            loc = locs[i].strip() if i < len(locs) else ""
            if marker and loc[:2] in ("* ", "- "):
                loc = loc[2:].strip()
            req.sources.append(Source(req.rid, i + 1, len(quotes), marker, quote, loc))
        if len(locs) != len(quotes):
            for s in req.sources:
                s.note = f"{len(quotes)} quotes but {len(locs)} locations"
        reqs.append(req)
    return cols, lang, reqs


def comment_text(src: Source, req: Requirement, words: dict) -> str:
    text = req.rid
    if src.n > 1:
        text += f" ({src.k}/{src.n})"
    if src.marker == "-":
        text += f", {words['qualifier']}"
    others = [o.location for o in req.sources if o is not src and o.location]
    if others:
        text += f" – {words['also']}: " + "; ".join(others)
    return text


def bookmark_name(rid: str, k: int) -> str:
    name = re.sub(r"\W", "_", rid, flags=re.ASCII)
    if not name[:1].isalpha():
        name = "R_" + name
    return f"{name}_{k}"[:40]


def next_found(req: Requirement, src: Source) -> Source | None:
    found = [s for s in req.sources if s.found]
    if len(found) < 2:
        return None
    return found[(found.index(src) + 1) % len(found)]


# ---------------------------------------------------------------- PDF

class PdfAnnotator:
    def __init__(self, path: Path):
        import pymupdf
        self.pymupdf = pymupdf
        self.doc = pymupdf.open(path)
        self.words = [p.get_text("words") for p in self.doc]
        if not any(self.words):
            sys.exit("The PDF has no text layer (scanned?). Run OCR on it first.")
        self.labels: dict[int, list[tuple[float, float]]] = {}

    def _stream(self, pages: list[int], column_order: bool):
        """Normalised text of the pages, with a (page, word index) reference per character.

        column_order puts the body text of all pages first and text in a right-hand margin
        column (e.g. review comments) after it, so passages crossing a page break stay together.
        """
        items = []
        for pg in pages:
            width = self.doc[pg].rect.width
            block_x: dict[int, float] = {}
            for w in self.words[pg]:
                block_x[w[5]] = min(block_x.get(w[5], w[0]), w[0])
            for wi, w in enumerate(self.words[pg]):
                margin = column_order and block_x[w[5]] > width * 0.6
                items.append(((margin, pg, w[5], w[6], w[7]) if column_order else (pg,), pg, wi, w[4]))
        if column_order:
            items.sort(key=lambda it: it[0])
        chars, refs = [], []
        for _, pg, wi, text in items:
            for c in text:
                for nc in norm(c):
                    chars.append(nc)
                    refs.append((pg, wi))
        return "".join(chars), refs

    def locate(self, src: Source) -> None:
        pat = pattern_for(src.quote)
        quote_c = "".join(pat.segments)
        if len(quote_c) < 4:
            src.status = "too short"
            return
        hint = [p for p in src.pages() if 0 <= p < len(self.doc)]
        near = sorted({q for p in hint for q in (p - 1, p, p + 1) if 0 <= q < len(self.doc)})
        ranges = [r for r in (near, list(range(len(self.doc)))) if r]
        for approximate in (False, True):
            if approximate and pat.kind != "text":
                break
            for pages in ranges:
                for column_order in (False, True):
                    text, refs = self._stream(pages, column_order)
                    spans = find_anchored(text, quote_c) if approximate else find_spans(text, pat)
                    if not spans:
                        continue
                    # several matches: prefer one starting on a hinted page
                    best = next((s for s in spans if refs[s[0]][0] in hint), spans[0])
                    words = sorted({refs[i] for i in range(best[0], best[1])})
                    src.anchor = words
                    distinct = {refs[s[0]] for s in spans}
                    if len(distinct) > 1:
                        src.status = "ambiguous"
                        src.note = f"{len(distinct)} matches, used page {words[0][0] + 1}"
                    else:
                        src.status = "approximate" if approximate else "found"
                    return
        src.status = "not found"

    def _line_rects(self, words):
        """Merge matched words into one rectangle per (page, block, line)."""
        lines: dict[tuple, object] = {}
        for pg, wi in words:
            w = self.words[pg][wi]
            key = (pg, w[5], w[6])
            r = self.pymupdf.Rect(w[:4])
            lines[key] = lines[key] | r if key in lines else r
        by_page: dict[int, list] = {}
        for (pg, _, _), r in lines.items():
            by_page.setdefault(pg, []).append(r)
        return by_page

    def _label_box(self, pg: int, y: float, height: float):
        page = self.doc[pg]
        xs = [w[0] for w in self.words[pg]] or [page.rect.width / 2]
        x1s = [w[2] for w in self.words[pg]] or [page.rect.width / 2]
        left, right = min(xs), page.rect.width - max(x1s)
        used = self.labels.setdefault(pg, [])
        for a, b in sorted(used):  # stack labels that would overlap
            if y < b + 1 and y + height > a - 1:
                y = b + 1
        used.append((y, y + height))
        if left >= 36 or left >= right:
            width = max(min(52, left - 6), 30)
            return self.pymupdf.Rect(4, y, 4 + width, y + height)
        width = max(min(52, right - 6), 30)
        return self.pymupdf.Rect(page.rect.width - 4 - width, y, page.rect.width - 4, y + height)

    def annotate(self, reqs: list[Requirement], words: dict, out: Path) -> None:
        pm = self.pymupdf
        toc = [[1, words["root"], 1, {"kind": pm.LINK_GOTO, "page": 0, "to": pm.Point(0, 0), "collapse": False}]]
        for req in reqs:
            first = True
            for src in req.sources:
                if not src.found:
                    continue
                rects = self._line_rects(src.anchor)
                text = comment_text(src, req, words)
                color = (0.62, 0.85, 1) if src.marker == "-" else (1, 0.92, 0.23)
                for pg, rs in rects.items():
                    page = self.doc[pg]  # keep a reference: annotations die with their page object
                    a = page.add_highlight_annot(rs)
                    a.set_colors(stroke=color)
                    a.set_info(title=AUTHOR, content=text)
                    a.update()
                pg0 = min(rects)
                top = min(r.y0 for r in rects[pg0])
                src.anchor = (pg0, top)
                multi = sum(s.found for s in req.sources) > 1
                box = self._label_box(pg0, top - 1, 17 if multi else 10)
                page = self.doc[pg0]
                page.draw_rect(box, color=(0.2, 0.4, 0.8), fill=(0.93, 0.96, 1), width=0.5)
                label = f"{req.rid}\n{src.k}/{src.n} »" if multi else req.rid
                for size in (5.5, 5, 4.5, 4):
                    if page.insert_textbox(box + (2, 0.5, -1, 0), label, fontsize=size, color=(0.1, 0.2, 0.6)) >= 0:
                        break
                src.box = box
                dest = {"kind": pm.LINK_GOTO, "page": pg0, "to": pm.Point(0, max(top - 40, 0)), "collapse": True}
                if first:
                    toc.append([2, req.rid, pg0 + 1, dest])
                    first = False
                if multi:
                    toc.append([3, f"{src.k}/{src.n} {src.location}"[:120], pg0 + 1, dict(dest)])
        for req in reqs:  # jump links, now that every target position is known
            for src in req.sources:
                nxt = next_found(req, src) if src.found else None
                if nxt:
                    pg, top = nxt.anchor
                    self.doc[src.anchor[0]].insert_link({
                        "kind": pm.LINK_GOTO, "from": src.box, "page": pg,
                        "to": pm.Point(0, max(top - 40, 0)), "zoom": 0})
        if len(toc) > 1:
            self.doc.set_toc(toc + self.doc.get_toc(simple=False))
            self.doc.xref_set_key(self.doc.pdf_catalog(), "PageMode", "/UseOutlines")
        self.doc.save(out, garbage=3, deflate=True)

    def link(self, src: Source) -> str:
        return f"page={src.anchor[0] + 1}"


# ---------------------------------------------------------------- Word

class WordAnnotator:
    def __init__(self, path: Path):
        import docx
        from docx.enum.text import WD_COLOR_INDEX
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.text.run import Run
        self.qn, self.OxmlElement, self.Run, self.Color = qn, OxmlElement, Run, WD_COLOR_INDEX
        self.doc = docx.Document(path)
        self.body = self.doc.element.body
        self.marker_anchors: set[str] = set()
        ids = [int(b.get(qn("w:id"))) for b in self.body.iter(qn("w:bookmarkStart")) if b.get(qn("w:id"), "").isdigit()]
        self.next_bookmark_id = max(ids, default=0) + 1

    # --- document text as a stream of runs

    def _para(self, el):
        while el is not None and el.tag != self.qn("w:p"):
            el = el.getparent()
        return el

    def _skip(self, r, p) -> bool:
        el = r.getparent()
        while el is not p:
            if el.tag == self.qn("w:del"):
                return True
            if el.tag == self.qn("w:hyperlink") and el.get(self.qn("w:anchor")) in self.marker_anchors:
                return True
            el = el.getparent()
        return False

    def _units(self, r):
        """(child, text) for the run's content children, in order."""
        qn = self.qn
        out = []
        for c in r:
            if c.tag == qn("w:t"):
                out.append((c, c.text or ""))
            elif c.tag in (qn("w:tab"), qn("w:br"), qn("w:cr")):
                out.append((c, " "))
            elif c.tag == qn("w:noBreakHyphen"):
                out.append((c, "-"))
            elif c.tag != qn("w:rPr"):
                out.append((c, ""))
        return out

    def _stream(self):
        chars, refs = [], []
        for p in self.body.iter(self.qn("w:p")):
            for r in p.iter(self.qn("w:r")):
                if self._para(r.getparent()) is not p or self._skip(r, p):
                    continue
                text = "".join(t for _, t in self._units(r))
                for i, c in enumerate(text):
                    for nc in norm(c):
                        chars.append(nc)
                        refs.append((r, i))
        return "".join(chars), refs

    def locate_all(self, reqs: list[Requirement]) -> None:
        text, _ = self._stream()
        for req in reqs:
            for src in req.sources:
                if src.status == "external":
                    continue
                pat = pattern_for(src.quote)
                quote_c = "".join(pat.segments)
                if len(quote_c) < 4:
                    src.status = "too short"
                    continue
                spans, approximate = find_spans(text, pat), False
                if not spans and pat.kind == "text":
                    spans, approximate = find_anchored(text, quote_c), True
                if not spans:
                    src.status = "not found"
                    continue
                src.anchor = spans[0]  # compact-stream offsets; stable while the text is unchanged
                if len(spans) > 1:
                    src.status, src.note = "ambiguous", f"{len(spans)} matches, used the first"
                else:
                    src.status = "approximate" if approximate else "found"

    # --- editing

    def _split(self, r, off: int):
        """Split run r at text offset off; return the run that starts at off."""
        units = self._units(r)
        if off <= 0:
            return r
        if off >= sum(len(t) for _, t in units):
            return None
        right = deepcopy(r)
        pos = 0
        for (lc, t), rc in zip(units, [c for c in right if c.tag != self.qn("w:rPr")]):
            start, end = pos, pos + len(t)
            if end <= off:
                right.remove(rc)
            elif start >= off:
                r.remove(lc)
            else:  # w:t containing the split point
                k = off - start
                lc.text, rc.text = t[:k], t[k:]
                for el in (lc, rc):
                    el.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            pos = end
        r.addnext(right)
        return right

    def _top(self, el, p):
        while el.getparent() is not p:
            el = el.getparent()
        return el

    def _after_annotations(self, el):
        """Last of el's following siblings that belong to annotations of an earlier source
        on the same passage, so a new one ends after them instead of inside them."""
        qn = self.qn
        while True:
            nxt = el.getnext()
            if nxt is None:
                return el
            is_ours = (
                nxt.tag in (qn("w:bookmarkEnd"), qn("w:commentRangeEnd"))
                or (nxt.tag == qn("w:hyperlink") and nxt.get(qn("w:anchor")) in self.marker_anchors)
                or (nxt.tag == qn("w:r") and nxt.find(qn("w:commentReference")) is not None)
            )
            if not is_ours:
                return el
            el = nxt

    def _link(self, text: str, anchor: str):
        qn, E = self.qn, self.OxmlElement
        h = E("w:hyperlink")
        h.set(qn("w:anchor"), anchor)
        h.set(qn("w:history"), "1")
        r, rpr = E("w:r"), E("w:rPr")
        for tag, val in (("w:color", "0563C1"), ("w:sz", "16"), ("w:u", "single")):
            el = E(tag)
            el.set(qn("w:val"), val)
            rpr.append(el)
        t = E("w:t")
        t.text = text
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.append(rpr)
        r.append(t)
        h.append(r)
        return h

    def annotate(self, reqs: list[Requirement], words: dict, out: Path) -> None:
        qn, E = self.qn, self.OxmlElement
        for req in reqs:
            multi = sum(s.found for s in req.sources) > 1
            for src in req.sources:
                if not src.found:
                    continue
                s, e = src.anchor
                _, refs = self._stream()
                (r_s, off_s), (r_e, off_e) = refs[s], refs[e - 1]
                self._split(r_e, off_e + 1)
                first = self._split(r_s, off_s)
                if r_e is r_s:
                    r_e = first
                runs, inside = [], False
                for p in self.body.iter(qn("w:p")):
                    for r in p.iter(qn("w:r")):
                        if self._para(r.getparent()) is not p or self._skip(r, p):
                            continue
                        inside = inside or r is first
                        if inside:
                            runs.append(r)
                        if r is r_e:
                            break
                    if runs and runs[-1] is r_e:
                        break
                color = self.Color.TURQUOISE if src.marker == "-" else self.Color.YELLOW
                for r in runs:
                    self.Run(r, None).font.highlight_color = color
                p_first, p_last = self._para(first), self._para(r_e)
                name = bookmark_name(req.rid, src.k)
                bs, be = E("w:bookmarkStart"), E("w:bookmarkEnd")
                bs.set(qn("w:id"), str(self.next_bookmark_id))
                bs.set(qn("w:name"), name)
                be.set(qn("w:id"), str(self.next_bookmark_id))
                self.next_bookmark_id += 1
                self._top(first, p_first).addprevious(bs)
                self._top(r_e, p_last).addnext(be)
                nxt = next_found(req, src)
                label = f" [{req.rid} {src.k}/{src.n} »]" if multi else f" [{req.rid}]"
                target = bookmark_name(req.rid, nxt.k) if nxt else name
                # the marker goes after markers of earlier sources on the same passage
                self._after_annotations(be).addnext(self._link(label, target))
                self.marker_anchors.add(target)
                self.doc.add_comment([self.Run(first, None), self.Run(r_e, None)],
                                     text=comment_text(src, req, words), author=AUTHOR, initials="ER")
                src.anchor = name
        self.doc.save(out)

    def link(self, src: Source) -> str:
        return src.anchor


# ---------------------------------------------------------------- main

def main() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    source, xlsx = Path(sys.argv[1]), Path(sys.argv[2])
    kind = source.suffix.lower()
    if kind not in (".pdf", ".docx"):
        sys.exit(f"Unsupported source format {kind!r}: use .pdf or .docx (save .doc files as .docx first).")
    out = xlsx.parent / f"{source.stem} - Annotated{kind}"

    wb = openpyxl.load_workbook(xlsx)
    ws = wb.active
    cols, lang, reqs = read_requirements(ws)
    words = LANG_WORDS[lang]
    sources = [s for r in reqs for s in r.sources]
    for s in sources:
        if s.location.lower().startswith(("http://", "https://")):
            s.status = "external"

    if kind == ".pdf":
        annotator = PdfAnnotator(source)
        for s in sources:
            if s.status != "external":
                annotator.locate(s)
    else:
        annotator = WordAnnotator(source)
        annotator.locate_all(reqs)
    annotator.annotate(reqs, words, out)

    for req in reqs:
        first = next((s for s in req.sources if s.found and s.marker in ("*", "")), None) \
            or next((s for s in req.sources if s.found), None)
        cell = ws.cell(req.row, cols["loc"])
        cell.hyperlink = None
        if first:
            cell.hyperlink = Hyperlink(ref=cell.coordinate, target=out.name, location=annotator.link(first))
            font = copy(cell.font)
            font.color, font.underline = "0563C1", "single"
            cell.font = font
    try:
        wb.save(xlsx)
    except PermissionError:
        sys.exit(f"Annotated document written, but {xlsx.name} is open in Excel: close it and run again.")

    counts = {}
    for s in sources:
        counts[s.status] = counts.get(s.status, 0) + 1
    print(f"Annotated: {out.name}")
    print(f"Sources: {len(sources)} | " + " | ".join(f"{k}: {v}" for k, v in sorted(counts.items())))
    problems = [s for s in sources if s.status in ("not found", "too short", "ambiguous", "approximate") or s.note]
    if problems:
        print("\nCheck these sources (fix the quote in the Excel file and run again):")
        for s in problems:
            note = f" ({s.note})" if s.note else ""
            print(f"- {s.req} source {s.k}/{s.n} [{s.status}]{note}: {s.location}\n    \"{s.quote[:150]}\"")


if __name__ == "__main__":
    main()
