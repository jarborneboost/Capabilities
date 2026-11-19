"""Annotating a PDF end to end, on a generated document (see conftest.pdf_case)."""
import openpyxl
import pymupdf
import pytest

SIMPLE_ROWS = [("REQ-001", '"Some text."', "page 1")]


@pytest.fixture
def run(ann, pdf_case):
    pdf, xlsx = pdf_case
    out, reqs = ann.annotate_files(pdf, xlsx)
    return out, {r.rid: r for r in reqs}, xlsx


def _status(req, k=0):
    return req.sources[k].match.status


def _span_text(out, src) -> str:
    """The text of the words a source matched, in reading order."""
    doc = pymupdf.open(out)
    words = [doc[p].get_text("words") for p in range(len(doc))]
    return " ".join(words[p][w][4] for p, w in src.match.span.words)


def test_statuses(ann, run):
    _, reqs, _ = run
    S = ann.Status
    assert _status(reqs["REQ-001"]) is S.FOUND
    assert _status(reqs["REQ-002"], 0) is S.FOUND  # crosses the page break, comment column in between
    assert _status(reqs["REQ-002"], 1) is S.FOUND  # quoted from the margin comment
    assert _status(reqs["REQ-003"]) is S.FOUND  # table row
    assert _status(reqs["REQ-004"]) is S.FOUND  # part of a margin comment
    assert _status(reqs["REQ-005"]) is S.FOUND  # on pages 1 and 3, but the stated page decides
    assert _status(reqs["REQ-006"]) is S.NOT_FOUND
    assert _status(reqs["REQ-007"]) is S.TOO_SHORT
    assert _status(reqs["REQ-008"]) is S.AMBIGUOUS  # on pages 1 and 3, no page stated


def test_matched_words_are_the_quote(ann, run):
    out, reqs, _ = run
    for rid in ("REQ-001", "REQ-002", "REQ-004"):
        for src in reqs[rid].sources:
            assert ann.compact(src.quote) in ann.compact(_span_text(out, src)), rid
    assert ann.compact(_span_text(out, reqs["REQ-003"].sources[0])) == ann.compact("Reminder 14 38,00")


def test_stated_page_decides_between_duplicates(run):
    _, reqs, _ = run
    assert reqs["REQ-005"].sources[0].placed.link == "page=3"
    ambiguous = reqs["REQ-008"].sources[0]
    assert ambiguous.placed.link == "page=1"  # no page stated: the first match
    assert ambiguous.match.note == "2 matches, used page 1"


def test_page_break_passage_highlighted_on_both_pages(run):
    out, reqs, _ = run
    doc = pymupdf.open(out)
    pages = {p for p, _ in reqs["REQ-002"].sources[0].match.span.words}
    assert pages == {0, 1}
    for p in pages:
        contents = [a.info["content"] for a in doc[p].annots()]
        assert "REQ-002 (1/2) – also: Section 1 Invoicing, page 2" in contents


def test_highlight_colours(ann, run):
    out, _, _ = run
    doc = pymupdf.open(out)
    colours = {a.info["content"].split(" ")[0] + ("-q" if "qualifier" in a.info["content"] else ""):
               tuple(round(c, 2) for c in a.colors["stroke"]) for p in doc for a in p.annots()}
    assert colours["REQ-001"] == tuple(round(c, 2) for c in ann.PDF_HIGHLIGHT["*"])
    assert colours["REQ-002-q"] == tuple(round(c, 2) for c in ann.PDF_HIGHLIGHT["-"])


def test_labels_are_searchable_text(run):
    out, reqs, _ = run
    doc = pymupdf.open(out)
    for rid in ("REQ-001", "REQ-003", "REQ-004", "REQ-005"):
        page = reqs[rid].sources[0].placed.jump_page
        assert doc[page].search_for(rid), rid
    assert not doc[0].search_for("REQ-006")  # not found: no label


def test_jump_links_cycle(run):
    out, reqs, _ = run
    doc = pymupdf.open(out)
    first, second = reqs["REQ-002"].sources
    for src, nxt in ((first, second), (second, first)):
        links = [l for l in doc[src.placed.jump_page].get_links()
                 if pymupdf.Rect(l["from"]).intersects(src.placed.label_box)]
        assert [l["page"] for l in links] == [nxt.placed.jump_page]
    assert not [l for l in doc[0].get_links() if l["from"].intersects(reqs["REQ-001"].sources[0].placed.label_box)]


def test_bookmarks_panel(ann, run):
    out, _, _ = run
    doc = pymupdf.open(out)
    toc = [(lvl, title) for lvl, title, _ in doc.get_toc()]
    assert toc[0] == (1, ann.LANG_WORDS["en"]["root"])
    assert (2, "REQ-001") in toc and (2, "REQ-006") not in toc
    assert (3, "1/2 Section 1 Invoicing, page 1-2") in toc
    assert doc.xref_get_key(doc.pdf_catalog(), "PageMode") == ("name", "/UseOutlines")


def test_excel_links(run):
    out, reqs, xlsx = run
    ws = openpyxl.load_workbook(xlsx).active
    links = {ws.cell(r.row, 1).value: ws.cell(r.row, 3).hyperlink for r in reqs.values()}
    assert (links["REQ-001"].target, links["REQ-001"].location) == (out.name, "page=1")
    assert links["REQ-005"].location == "page=3"
    assert links["REQ-006"] is None and links["REQ-007"] is None


def test_rerun_gives_the_same_result(ann, run, pdf_case):
    out, _, xlsx = run
    first = [(p.number, a.info["content"]) for p in pymupdf.open(out) for a in p.annots()]
    ann.annotate_files(pdf_case[0], xlsx)
    second = [(p.number, a.info["content"]) for p in pymupdf.open(out) for a in p.annots()]
    assert first == second
    ws = openpyxl.load_workbook(xlsx).active
    assert sum(1 for row in ws.iter_rows() for c in row if c.hyperlink) == 6  # REQ-001..005 and REQ-008


def test_pdf_without_text_layer(ann, tmp_path, make_xlsx):
    blank = tmp_path / "scan.pdf"
    doc = pymupdf.open()
    doc.new_page()
    doc.save(blank)
    with pytest.raises(ann.AnnotateError, match="no text layer"):
        ann.annotate_files(blank, make_xlsx(tmp_path / "r.xlsx", SIMPLE_ROWS))


def test_unsupported_format(ann, tmp_path, make_xlsx):
    with pytest.raises(ann.AnnotateError, match="Unsupported source format"):
        ann.annotate_files(tmp_path / "old.doc", make_xlsx(tmp_path / "r.xlsx", SIMPLE_ROWS))
