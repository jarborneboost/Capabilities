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

import enum
import re
import sys
from copy import copy, deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

import docx
import openpyxl
import pymupdf
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.run import Run
from openpyxl.worksheet.hyperlink import Hyperlink

# ================================================================= constants

AUTHOR = "extract-requirements"  # author shown on highlights and comments
AUTHOR_INITIALS = "ER"

# --- matching
MIN_QUOTE_CHARS = 4  # shorter quotes match too many places to be useful
MAX_MATCHES = 5  # enough to tell "found" from "ambiguous"; more only costs time
TABLE_MAX_GAP = 3  # chars allowed between the cell values of a table row, e.g. a row number
ELLIPSIS_MAX_GAP = 3000  # chars an omission marked with "..." may skip, about one page
ANCHOR_LEN = 30  # chars matched at the start and at the end of a long quote as a fallback
ANCHOR_SLACK_RATIO = 1.5  # an anchored match may be this much longer than the quote,
ANCHOR_SLACK_CHARS = 200  # plus this, for page headers/footers or bullets inside the passage

# --- PDF layout (points)
MARGIN_COLUMN_X_RATIO = 0.6  # blocks starting right of this share of the page width form a
#                              margin column (e.g. review comments), read after the body text
LABEL_EDGE = 4  # distance of a label from the page edge
LABEL_MIN_MARGIN = 36  # free left margin needed for labels there; otherwise the wider side
LABEL_MIN_WIDTH = 30
LABEL_MAX_WIDTH = 52
LABEL_HEIGHT_SINGLE = 10  # one line: the requirement ID
LABEL_HEIGHT_MULTI = 17  # two lines: the ID and "k/n »"
LABEL_GAP = 1  # vertical space between stacked labels
LABEL_FONT_SIZES = (5.5, 5, 4.5, 4)  # tried in order until the label text fits
JUMP_ABOVE = 40  # jumps land this far above a passage, so it isn't at the window's edge
TOC_TITLE_MAX_LEN = 120  # long Location texts make the bookmarks panel unreadable

# --- colours
PDF_HIGHLIGHT = {"*": (1, 0.92, 0.23), "": (1, 0.92, 0.23), "-": (0.62, 0.85, 1)}  # yellow / light blue
WORD_HIGHLIGHT = {"*": WD_COLOR_INDEX.YELLOW, "": WD_COLOR_INDEX.YELLOW, "-": WD_COLOR_INDEX.TURQUOISE}
PDF_LABEL_BORDER, PDF_LABEL_FILL, PDF_LABEL_TEXT = (0.2, 0.4, 0.8), (0.93, 0.96, 1), (0.1, 0.2, 0.6)
LINK_COLOR = "0563C1"  # Office's default hyperlink blue

# --- Word
BOOKMARK_MAX_LEN = 40  # Word drops bookmarks with longer names
MARKER_HALF_POINTS = "16"  # marker font size (8 pt) in Word's half-point unit
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# --- Excel and wording
HEADERS = {
    "id": {"id"},
    "text": {"original text", "opprinnelig tekst"},
    "loc": {"location", "plassering"},
}
NORWEGIAN_LOCATION_HEADER = "plassering"
LANG_WORDS = {
    "en": {"also": "also", "qualifier": "qualifier", "root": "Requirements"},
    "no": {"also": "også", "qualifier": "presisering", "root": "Krav"},
}
PAGE_RE = re.compile(r"\b(?:page|side)\s+(\d+)(?:\s*[-–]\s*(\d+))?", re.I)
ELLIPSIS_RE = re.compile(r"\[\.\.\.\]|\.\.\.|…")
EXTERNAL_PREFIXES = ("http://", "https://")


class AnnotateError(Exception):
    """A problem with the input that the user has to fix; reported without a traceback."""


# ================================================================= data model

class Status(enum.Enum):
    FOUND = "found"
    APPROXIMATE = "approximate"  # only the start and end of a long quote matched
    AMBIGUOUS = "ambiguous"  # several matches; the one on the stated page (or the first) is used
    NOT_FOUND = "not found"
    EXTERNAL = "external"  # cited from a link, not from the document
    TOO_SHORT = "too short"

    @property
    def found(self) -> bool:
        return self in (Status.FOUND, Status.APPROXIMATE, Status.AMBIGUOUS)


@dataclass(frozen=True)
class PdfSpan:
    words: tuple[tuple[int, int], ...]  # (page index, word index), sorted


@dataclass(frozen=True)
class WordSpan:
    start: int  # offsets in the normalised document text; end is exclusive
    end: int


@dataclass
class Match:
    """Result of locating a source: where it is, or why it isn't."""
    status: Status
    note: str = ""
    span: PdfSpan | WordSpan | None = None


@dataclass
class Placed:
    """Result of annotating a source: where links to it point."""
    link: str  # Excel sub-address: "page=N" (PDF) or a bookmark name (Word)
    jump_page: int = 0  # PDF only: target of jump links
    jump_y: float = 0.0
    label_box: pymupdf.Rect | None = None  # PDF only: the clickable margin label


@dataclass(eq=False)
class Source:
    """One quoted passage of a requirement, as read from the Excel file."""
    req: str
    k: int  # 1-based position among the requirement's sources
    n: int
    marker: str  # "*" obligation, "-" qualifier, "" single source
    quote: str
    location: str
    warning: str = ""  # problem with the Excel row itself
    match: Match | None = None
    placed: Placed | None = None

    @property
    def found(self) -> bool:
        return self.match is not None and self.match.status.found

    def page_hint(self) -> list[int]:
        """0-based pages named in the location, e.g. "page 28-29" -> [27, 28]."""
        m = PAGE_RE.search(self.location)
        if not m:
            return []
        first = int(m.group(1))
        return list(range(first - 1, int(m.group(2) or first)))


@dataclass
class Requirement:
    rid: str
    row: int  # Excel row
    sources: list[Source] = field(default_factory=list)

    def found_sources(self) -> list[Source]:
        return [s for s in self.sources if s.found]

    def next_found(self, src: Source) -> Source | None:
        """The source a jump link from src leads to, cycling; None with fewer than two."""
        found = self.found_sources()
        if len(found) < 2:
            return None
        return found[(found.index(src) + 1) % len(found)]

    def link_source(self) -> Source | None:
        """The source the Location cell links to: the first found obligation, else any found."""
        found = self.found_sources()
        return next((s for s in found if s.marker in ("*", "")), found[0] if found else None)


# ================================================================= text matching

_TRANS = str.maketrans({
    "‘": "'", "’": "'", "‚": "'", "“": '"', "”": '"', "„": '"',
    "«": '"', "»": '"', "–": "-", "—": "-", "−": "-", "‑": "-",
    "­": "", "•": "", "": "", "▪": "", "●": "", "…": "...",
})


def norm(c: str) -> str:
    """Normalised, whitespace-free form of one character (may be empty or several chars).

    Typographic quotes and dashes become plain ones, soft hyphens and bullets disappear,
    case is ignored: quotes copied from a document match its text despite such differences.
    """
    return "".join(x for x in c.translate(_TRANS).lower() if not x.isspace())


def compact(s: str) -> str:
    return "".join(norm(c) for c in s)


@dataclass
class Pattern:
    segments: list[str]  # normalised parts that must occur in order
    max_gap: int  # chars allowed between two segments
    kind: str  # "text", "table" or "ellipsis"

    @property
    def length(self) -> int:
        return sum(len(s) for s in self.segments)


def pattern_for(quote: str) -> Pattern:
    if "|" in quote:  # table row: cell values separated by " | "
        return Pattern([s for s in map(compact, quote.split("|")) if s], TABLE_MAX_GAP, "table")
    parts = ELLIPSIS_RE.split(quote)
    if len(parts) > 1:
        return Pattern([s for s in map(compact, parts) if s], ELLIPSIS_MAX_GAP, "ellipsis")
    return Pattern([compact(quote)], 0, "text")


def find_spans(text: str, pat: Pattern) -> list[tuple[int, int]]:
    """(start, end) spans in text where the segments occur in order within max_gap."""
    spans: list[tuple[int, int]] = []
    if not pat.segments:
        return spans
    pos = 0
    while len(spans) < MAX_MATCHES:
        start = text.find(pat.segments[0], pos)
        if start < 0:
            break
        end = start + len(pat.segments[0])
        for seg in pat.segments[1:]:
            j = text.find(seg, end)
            if j < 0 or j - end > pat.max_gap:
                break
            end = j + len(seg)
        else:
            spans.append((start, end))
        pos = start + 1
    return spans


def find_anchored(text: str, quote_c: str) -> list[tuple[int, int]]:
    """Fallback for long quotes: the first and last ANCHOR_LEN chars, close enough together."""
    spans: list[tuple[int, int]] = []
    if len(quote_c) < 3 * ANCHOR_LEN:
        return spans
    head, tail = quote_c[:ANCHOR_LEN], quote_c[-ANCHOR_LEN:]
    max_len = len(quote_c) * ANCHOR_SLACK_RATIO + ANCHOR_SLACK_CHARS
    pos = 0
    while len(spans) < MAX_MATCHES:
        start = text.find(head, pos)
        if start < 0:
            break
        t = text.find(tail, start + ANCHOR_LEN)
        if t >= 0 and t + ANCHOR_LEN - start <= max_len:
            spans.append((start, t + ANCHOR_LEN))
        pos = start + 1
    return spans


def search(text: str, quote: str) -> tuple[list[tuple[int, int]], Status] | Status:
    """Spans of quote in text, and FOUND or APPROXIMATE; or TOO_SHORT / NOT_FOUND."""
    pat = pattern_for(quote)
    if pat.length < MIN_QUOTE_CHARS:
        return Status.TOO_SHORT
    spans = find_spans(text, pat)
    if spans:
        return spans, Status.FOUND
    if pat.kind == "text":
        spans = find_anchored(text, "".join(pat.segments))
        if spans:
            return spans, Status.APPROXIMATE
    return Status.NOT_FOUND


# ================================================================= Excel

def strip_quote(s: str) -> str:
    s = s.strip()
    if len(s) >= 2 and s[0] in "\"“«„'" and s[-1] in "\"”»“'":
        s = s[1:-1].strip()
    return s


def _lines(value) -> list[str]:
    return [line.strip() for line in str(value or "").splitlines() if line.strip()]


def read_requirements(ws) -> tuple[dict[str, int], str, list[Requirement]]:
    """Column numbers, output language ("en"/"no") and requirements of the sheet."""
    header = {str(c.value or "").strip().lower(): c.column for c in ws[1]}
    cols: dict[str, int] = {}
    for key, names in HEADERS.items():
        col = next((c for name, c in header.items() if name in names), None)
        if col is None:
            raise AnnotateError(f"Column not found in the Excel header: one of {sorted(names)}")
        cols[key] = col
    lang = "no" if NORWEGIAN_LOCATION_HEADER in header else "en"

    reqs = []
    for row in range(2, ws.max_row + 1):
        rid = ws.cell(row, cols["id"]).value
        if not rid:
            continue
        req = Requirement(str(rid).strip(), row)
        quotes, locs = _lines(ws.cell(row, cols["text"]).value), _lines(ws.cell(row, cols["loc"]).value)
        warning = f"{len(quotes)} quotes but {len(locs)} locations" if len(quotes) != len(locs) else ""
        for i, line in enumerate(quotes):
            marker = line[0] if line[:2] in ("* ", "- ") else ""
            loc = locs[i] if i < len(locs) else ""
            if loc[:2] in ("* ", "- "):
                loc = loc[2:].strip()
            src = Source(req.rid, i + 1, len(quotes), marker, strip_quote(line[2:] if marker else line),
                         loc, warning)
            if loc.lower().startswith(EXTERNAL_PREFIXES):
                src.match = Match(Status.EXTERNAL)
            req.sources.append(src)
        reqs.append(req)
    return cols, lang, reqs


def write_links(ws, loc_col: int, reqs: list[Requirement], target: str) -> None:
    """Link each Location cell to its requirement's link source; remove stale links."""
    for req in reqs:
        cell = ws.cell(req.row, loc_col)
        src = req.link_source()
        font = copy(cell.font)
        if src and src.placed:
            cell.hyperlink = Hyperlink(ref=cell.coordinate, target=target, location=src.placed.link)
            font.color, font.underline = LINK_COLOR, "single"
        else:
            cell.hyperlink = None
            if font.color is not None and str(font.color.rgb).endswith(LINK_COLOR):
                font.color, font.underline = None, None
        cell.font = font


# ================================================================= wording

def comment_text(src: Source, req: Requirement, words: dict[str, str]) -> str:
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
    """Word bookmark name: letters, digits and "_" only, starting with a letter."""
    name = re.sub(r"\W", "_", rid, flags=re.ASCII)
    if not name[:1].isalpha():
        name = "R_" + name
    suffix = f"_{k}"
    return name[: BOOKMARK_MAX_LEN - len(suffix)] + suffix


# ================================================================= annotators

class Annotator(Protocol):
    def locate(self, sources: list[Source]) -> None:
        """Set .match on every source that has none yet."""

    def annotate(self, reqs: list[Requirement], words: dict[str, str], out: Path) -> None:
        """Mark every found source, set its .placed, and save the annotated copy to out."""


# ----------------------------------------------------------------- PDF

class PdfAnnotator:
    def __init__(self, path: Path):
        self.doc = pymupdf.open(path)
        self.words = [page.get_text("words") for page in self.doc]
        if not any(self.words):
            raise AnnotateError("The PDF has no text layer (scanned?). Run OCR on it first.")
        self._streams: dict[tuple, tuple[str, list[tuple[int, int]]]] = {}
        self._labels: dict[int, list[tuple[float, float]]] = {}  # used label y-ranges per page

    # --- locating

    def _stream(self, pages: tuple[int, ...], column_order: bool) -> tuple[str, list[tuple[int, int]]]:
        """Normalised text of the pages with a (page, word index) per character.

        column_order puts the body text of all pages first and text in a right-hand margin
        column after it, so passages crossing a page break stay together.
        """
        key = (pages, column_order)
        if key not in self._streams:
            items = []
            for pg in pages:
                limit = self.doc[pg].rect.width * MARGIN_COLUMN_X_RATIO
                block_x: dict[int, float] = {}
                for w in self.words[pg]:
                    block_x[w[5]] = min(block_x.get(w[5], w[0]), w[0])
                for wi, w in enumerate(self.words[pg]):
                    order = (block_x[w[5]] > limit, pg, w[5], w[6], w[7]) if column_order else (pg,)
                    items.append((order, pg, wi, w[4]))
            if column_order:
                items.sort(key=lambda item: item[0])
            chars, refs = [], []
            for _, pg, wi, text in items:
                for c in text:
                    for nc in norm(c):
                        chars.append(nc)
                        refs.append((pg, wi))
            self._streams[key] = ("".join(chars), refs)
        return self._streams[key]

    def _locate(self, src: Source) -> Match:
        hint = [p for p in src.page_hint() if 0 <= p < len(self.doc)]
        near = tuple(sorted({q for p in hint for q in (p - 1, p, p + 1) if 0 <= q < len(self.doc)}))
        ranges = [r for r in (near, tuple(range(len(self.doc)))) if r]
        best: tuple[list[tuple[int, int]], list, Status] | None = None
        for pages in ranges:
            for column_order in (False, True):
                text, refs = self._stream(pages, column_order)
                result = search(text, src.quote)
                if result is Status.TOO_SHORT:
                    return Match(Status.TOO_SHORT)
                if isinstance(result, Status):
                    continue
                spans, status = result
                if status is Status.FOUND:  # an exact match beats any approximate one
                    return self._match(spans, refs, status, hint)
                best = best or (spans, refs, status)
        if best:
            return self._match(*best, hint)
        return Match(Status.NOT_FOUND)

    def _match(self, spans, refs, status: Status, hint: list[int]) -> Match:
        chosen = next((s for s in spans if refs[s[0]][0] in hint), spans[0])
        words = tuple(sorted({refs[i] for i in range(*chosen)}))
        starts = {refs[s[0]] for s in spans}
        if len(starts) > 1:
            return Match(Status.AMBIGUOUS, f"{len(starts)} matches, used page {words[0][0] + 1}", PdfSpan(words))
        return Match(status, span=PdfSpan(words))

    def locate(self, sources: list[Source]) -> None:
        for src in sources:
            if src.match is None:
                src.match = self._locate(src)

    # --- annotating

    def _line_rects(self, span: PdfSpan) -> dict[int, list[pymupdf.Rect]]:
        """One rectangle per (page, block, line) of the matched words, grouped by page."""
        lines: dict[tuple[int, int, int], pymupdf.Rect] = {}
        for pg, wi in span.words:
            w = self.words[pg][wi]
            key, rect = (pg, w[5], w[6]), pymupdf.Rect(w[:4])
            lines[key] = lines[key] | rect if key in lines else rect
        by_page: dict[int, list[pymupdf.Rect]] = {}
        for (pg, _, _), rect in lines.items():
            by_page.setdefault(pg, []).append(rect)
        return by_page

    def _label_box(self, pg: int, y: float, height: float) -> pymupdf.Rect:
        """A free spot for a label at height y in the wider margin, below earlier labels."""
        width = self.doc[pg].rect.width
        left = min((w[0] for w in self.words[pg]), default=width / 2)
        right = width - max((w[2] for w in self.words[pg]), default=width / 2)
        used = self._labels.setdefault(pg, [])
        for top, bottom in sorted(used):
            if y < bottom + LABEL_GAP and y + height > top - LABEL_GAP:
                y = bottom + LABEL_GAP
        used.append((y, y + height))
        if left >= LABEL_MIN_MARGIN or left >= right:
            w = max(min(LABEL_MAX_WIDTH, left - 2 * LABEL_EDGE), LABEL_MIN_WIDTH)
            return pymupdf.Rect(LABEL_EDGE, y, LABEL_EDGE + w, y + height)
        w = max(min(LABEL_MAX_WIDTH, right - 2 * LABEL_EDGE), LABEL_MIN_WIDTH)
        return pymupdf.Rect(width - LABEL_EDGE - w, y, width - LABEL_EDGE, y + height)

    def _mark(self, src: Source, req: Requirement, words: dict[str, str]) -> None:
        rects = self._line_rects(src.match.span)
        text = comment_text(src, req, words)
        for pg, page_rects in rects.items():
            page = self.doc[pg]  # keep a reference: annotations die with their page object
            annot = page.add_highlight_annot(page_rects)
            annot.set_colors(stroke=PDF_HIGHLIGHT[src.marker])
            annot.set_info(title=AUTHOR, content=text)
            annot.update()
        pg = min(rects)
        top = min(r.y0 for r in rects[pg])
        multi = len(req.found_sources()) > 1
        box = self._label_box(pg, top - 1, LABEL_HEIGHT_MULTI if multi else LABEL_HEIGHT_SINGLE)
        page = self.doc[pg]
        page.draw_rect(box, color=PDF_LABEL_BORDER, fill=PDF_LABEL_FILL, width=0.5)
        label = f"{req.rid}\n{src.k}/{src.n} »" if multi else req.rid
        for size in LABEL_FONT_SIZES:
            if page.insert_textbox(box + (2, 0.5, -1, 0), label, fontsize=size, color=PDF_LABEL_TEXT) >= 0:
                break
        src.placed = Placed(link=f"page={pg + 1}", jump_page=pg, jump_y=top, label_box=box)

    def _dest(self, placed: Placed, collapse: bool = True) -> dict:
        return {"kind": pymupdf.LINK_GOTO, "page": placed.jump_page,
                "to": pymupdf.Point(0, max(placed.jump_y - JUMP_ABOVE, 0)), "collapse": collapse}

    def annotate(self, reqs: list[Requirement], words: dict[str, str], out: Path) -> None:
        for req in reqs:
            for src in req.found_sources():
                self._mark(src, req, words)
        toc: list[list] = []
        for req in reqs:
            found = req.found_sources()
            if not found:
                continue
            toc.append([2, req.rid, found[0].placed.jump_page + 1, self._dest(found[0].placed)])
            for src in found:
                if len(found) > 1:
                    title = f"{src.k}/{src.n} {src.location}"[:TOC_TITLE_MAX_LEN]
                    toc.append([3, title, src.placed.jump_page + 1, self._dest(src.placed)])
                nxt = req.next_found(src)
                if nxt:
                    link = {"kind": pymupdf.LINK_GOTO, "from": src.placed.label_box, "zoom": 0,
                            "page": nxt.placed.jump_page, "to": self._dest(nxt.placed)["to"]}
                    self.doc[src.placed.jump_page].insert_link(link)
        if toc:
            root = [1, words["root"], 1, {"kind": pymupdf.LINK_GOTO, "page": 0, "to": pymupdf.Point(0, 0),
                                          "collapse": False}]
            self.doc.set_toc([root] + toc + self.doc.get_toc(simple=False))
            self.doc.xref_set_key(self.doc.pdf_catalog(), "PageMode", "/UseOutlines")
        self.doc.save(out, garbage=3, deflate=True)


# ----------------------------------------------------------------- Word

class WordAnnotator:
    """Works on the document XML: Word stores a sentence as several runs (pieces of text
    with one formatting), so a quote is found in the joined text and its runs are split
    where the quote starts and ends."""

    def __init__(self, path: Path):
        self.doc = docx.Document(path)
        self.body = self.doc.element.body
        self.marker_anchors: set[str] = set()  # anchors of our marker links, kept out of the text
        ids = [int(b.get(qn("w:id"))) for b in self.body.iter(qn("w:bookmarkStart"))
               if b.get(qn("w:id"), "").isdigit()]
        self.next_bookmark_id = max(ids, default=0) + 1

    # --- the document text

    @staticmethod
    def _para(el):
        while el is not None and el.tag != qn("w:p"):
            el = el.getparent()
        return el

    def _hidden(self, r, p) -> bool:
        """Runs that are not document text: deleted (tracked change) or our own markers."""
        el = r.getparent()
        while el is not p:
            if el.tag == qn("w:del"):
                return True
            if el.tag == qn("w:hyperlink") and el.get(qn("w:anchor")) in self.marker_anchors:
                return True
            el = el.getparent()
        return False

    @staticmethod
    def _units(r) -> list[tuple[object, str]]:
        """(child element, text) for the run's content, in order; tabs and breaks count as spaces."""
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

    def _runs(self) -> list:
        """All text runs in document order, including tables and text boxes."""
        runs = []
        for p in self.body.iter(qn("w:p")):
            runs.extend(r for r in p.iter(qn("w:r")) if self._para(r.getparent()) is p and not self._hidden(r, p))
        return runs

    def _stream(self) -> tuple[str, list[tuple[object, int]], list]:
        """Normalised text, a (run, offset in run) per character, and the runs in order."""
        runs = self._runs()
        chars, refs = [], []
        for r in runs:
            for i, c in enumerate("".join(t for _, t in self._units(r))):
                for nc in norm(c):
                    chars.append(nc)
                    refs.append((r, i))
        return "".join(chars), refs, runs

    def locate(self, sources: list[Source]) -> None:
        text, _, _ = self._stream()
        for src in sources:
            if src.match is not None:
                continue
            result = search(text, src.quote)
            if isinstance(result, Status):
                src.match = Match(result)
                continue
            spans, status = result
            note = ""
            if len(spans) > 1:
                status, note = Status.AMBIGUOUS, f"{len(spans)} matches, used the first"
            src.match = Match(status, note, WordSpan(*spans[0]))

    # --- editing

    def _split(self, r, off: int) -> None:
        """Split run r so that a new run starts at text offset off (no-op at the run's edges)."""
        units = self._units(r)
        if off <= 0 or off >= sum(len(t) for _, t in units):
            return
        right = deepcopy(r)
        pos = 0
        for (left_child, t), right_child in zip(units, [c for c in right if c.tag != qn("w:rPr")]):
            start, end = pos, pos + len(t)
            if end <= off:
                right.remove(right_child)
            elif start >= off:
                r.remove(left_child)
            else:  # the w:t containing the split point
                left_child.text, right_child.text = t[: off - start], t[off - start:]
                left_child.set(XML_SPACE, "preserve")
                right_child.set(XML_SPACE, "preserve")
            pos = end
        r.addnext(right)

    def _split_at_boundaries(self, spans: list[WordSpan]) -> None:
        """Split runs so every span starts and ends on a run boundary.

        Works from the end of the document backwards: splitting a run only adds a run after
        it, so the (run, offset) positions of earlier boundaries stay valid.
        """
        text, refs, runs = self._stream()
        order = {id(r): i for i, r in enumerate(runs)}
        cuts = set()
        for span in spans:
            r, off = refs[span.start]
            cuts.add((order[id(r)], off, r))
            r, off = refs[span.end - 1]
            cuts.add((order[id(r)], off + 1, r))
        for _, off, r in sorted(cuts, key=lambda c: (c[0], c[1]), reverse=True):
            self._split(r, off)

    @staticmethod
    def _top(el, p):
        """el's ancestor (or el itself) that is a direct child of paragraph p."""
        while el.getparent() is not p:
            el = el.getparent()
        return el

    def _after_annotations(self, el):
        """The last of el's following siblings that belong to annotations of an earlier source
        on the same passage, so a new marker comes after them instead of inside them."""
        while (nxt := el.getnext()) is not None and (
            nxt.tag in (qn("w:bookmarkEnd"), qn("w:commentRangeEnd"))
            or (nxt.tag == qn("w:hyperlink") and nxt.get(qn("w:anchor")) in self.marker_anchors)
            or (nxt.tag == qn("w:r") and nxt.find(qn("w:commentReference")) is not None)
        ):
            el = nxt
        return el

    @staticmethod
    def _link(text: str, anchor: str):
        """A small blue hyperlink to a bookmark in the same document."""
        h = OxmlElement("w:hyperlink")
        h.set(qn("w:anchor"), anchor)
        h.set(qn("w:history"), "1")
        r, rpr = OxmlElement("w:r"), OxmlElement("w:rPr")
        for tag, val in (("w:color", LINK_COLOR), ("w:sz", MARKER_HALF_POINTS), ("w:u", "single")):
            el = OxmlElement(tag)
            el.set(qn("w:val"), val)
            rpr.append(el)
        t = OxmlElement("w:t")
        t.text = text
        t.set(XML_SPACE, "preserve")
        r.append(rpr)
        r.append(t)
        h.append(r)
        return h

    def _bookmark(self, name: str, first, last) -> object:
        """Bookmark from the start of run first to the end of run last; returns its end element."""
        start, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
        start.set(qn("w:id"), str(self.next_bookmark_id))
        start.set(qn("w:name"), name)
        end.set(qn("w:id"), str(self.next_bookmark_id))
        self.next_bookmark_id += 1
        self._top(first, self._para(first)).addprevious(start)
        self._top(last, self._para(last)).addnext(end)
        return end

    def annotate(self, reqs: list[Requirement], words: dict[str, str], out: Path) -> None:
        found = [s for req in reqs for s in req.found_sources()]
        self._split_at_boundaries([s.match.span for s in found])
        _, refs, runs = self._stream()  # one rebuild: the annotations below add no text
        order = {id(r): i for i, r in enumerate(runs)}
        for req in reqs:
            multi = len(req.found_sources()) > 1
            for src in req.found_sources():
                span = src.match.span
                first, last = refs[span.start][0], refs[span.end - 1][0]
                passage = runs[order[id(first)]: order[id(last)] + 1]  # includes space-only runs
                for r in passage:
                    Run(r, None).font.highlight_color = WORD_HIGHLIGHT[src.marker]
                name = bookmark_name(req.rid, src.k)
                end = self._bookmark(name, first, last)
                nxt = req.next_found(src)
                target = bookmark_name(req.rid, nxt.k) if nxt else name
                label = f" [{req.rid} {src.k}/{src.n} »]" if multi else f" [{req.rid}]"
                self._after_annotations(end).addnext(self._link(label, target))
                self.marker_anchors.add(target)
                self.doc.add_comment([Run(first, None), Run(last, None)], text=comment_text(src, req, words),
                                     author=AUTHOR, initials=AUTHOR_INITIALS)
                src.placed = Placed(link=name)
        self.doc.save(out)


# ================================================================= running

def annotate_files(source: Path, xlsx: Path) -> tuple[Path, list[Requirement]]:
    """Annotate source with the sources in xlsx and link xlsx to it; returns the output path."""
    kind = source.suffix.lower()
    annotators = {".pdf": PdfAnnotator, ".docx": WordAnnotator}
    if kind not in annotators:
        raise AnnotateError(f"Unsupported source format {kind!r}: use .pdf or .docx "
                            "(save .doc files as .docx first).")
    out = xlsx.parent / f"{source.stem} - Annotated{kind}"
    wb = openpyxl.load_workbook(xlsx)
    ws = wb.active
    cols, lang, reqs = read_requirements(ws)
    annotator: Annotator = annotators[kind](source)
    annotator.locate([s for req in reqs for s in req.sources])
    annotator.annotate(reqs, LANG_WORDS[lang], out)
    write_links(ws, cols["loc"], reqs, out.name)
    try:
        wb.save(xlsx)
    except PermissionError:
        raise AnnotateError(f"Annotated document written, but {xlsx.name} is open in Excel: "
                            "close it and run again.") from None
    return out, reqs


def report(out: Path, reqs: list[Requirement]) -> str:
    sources = [s for req in reqs for s in req.sources]
    counts: dict[str, int] = {}
    for s in sources:
        counts[s.match.status.value] = counts.get(s.match.status.value, 0) + 1
    lines = [f"Annotated: {out.name}",
             f"Sources: {len(sources)} | " + " | ".join(f"{k}: {v}" for k, v in sorted(counts.items()))]
    attention = (Status.NOT_FOUND, Status.TOO_SHORT, Status.AMBIGUOUS, Status.APPROXIMATE)
    problems = [s for s in sources if s.match.status in attention or s.warning]
    if problems:
        lines.append("\nCheck these sources (fix the quote in the Excel file and run again):")
        for s in problems:
            notes = ", ".join(x for x in (s.match.note, s.warning) if x)
            notes = f" ({notes})" if notes else ""
            lines.append(f"- {s.req} source {s.k}/{s.n} [{s.match.status.value}]{notes}: {s.location}\n"
                         f"    \"{s.quote[:150]}\"")
    return "\n".join(lines)


def main(argv: list[str] | None = None) -> None:
    argv = sys.argv[1:] if argv is None else argv
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if len(argv) != 2:
        sys.exit(__doc__)
    try:
        out, reqs = annotate_files(Path(argv[0]), Path(argv[1]))
    except AnnotateError as e:
        sys.exit(str(e))
    print(report(out, reqs))


if __name__ == "__main__":
    main()
