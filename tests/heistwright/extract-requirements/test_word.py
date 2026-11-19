"""Annotating a Word document end to end, on a generated document (see conftest.word_case)."""
import docx
import openpyxl
import pytest
from docx.oxml.ns import qn

PASSAGES = {
    "REQ_001_1": "Fakturaen skal sendes innen 14 dager etter forfall.",
    "REQ_002_1": "Purringer skal inneholde følgende: saksnummer forfallsdato beløp",  # across paragraphs
    "REQ_002_2": "Gebyret skal være 38 kroner for både purring og inkassovarsel.",
    "REQ_003_1": "Inkassovarsel skal sendes etter 14 dager.",  # after the tracked change 10 -> 14
    "REQ_003_2": "Inkassovarsel 33 38,00 100,00",  # table row
    "REQ_003_3": "Gebyret skal være 38 kroner for både purring og inkassovarsel.",  # same passage as REQ_002_2
    "REQ_004_1": "Purring 14 38,00",  # table row, empty last cell
}


@pytest.fixture
def run(ann, word_case):
    source, xlsx = word_case
    out, reqs = ann.annotate_files(source, xlsx)
    return out, {r.rid: r for r in reqs}, xlsx


def _body(path):
    return docx.Document(path).element.body


def _marker_anchors(body) -> set[str]:
    return {h.get(qn("w:anchor")) for h in body.iter(qn("w:hyperlink"))}


def _visible_text(body, skip_hyperlinks: bool) -> str:
    """Document text (w:t), optionally without the text inside hyperlinks (our markers)."""
    out = []
    for t in body.iter(qn("w:t")):
        if skip_hyperlinks and any(a.tag == qn("w:hyperlink") for a in t.iterancestors()):
            continue
        out.append(t.text or "")
    return "".join(out)


def _ranges(body, start_tag, end_tag, key) -> dict[str, str]:
    """Text between start/end elements with the same w:id (bookmarks, comment ranges)."""
    names, open_, out = {}, {}, {}
    for el in body.iter():
        if el.tag == start_tag:
            names[el.get(qn("w:id"))] = key(el)
            open_[el.get(qn("w:id"))] = []
        elif el.tag == end_tag and el.get(qn("w:id")) in open_:
            out[names[el.get(qn("w:id"))]] = "".join(open_.pop(el.get(qn("w:id"))))
        elif el.tag == qn("w:t") and not any(a.tag == qn("w:hyperlink") for a in el.iterancestors()):
            for parts in open_.values():
                parts.append(el.text or "")
    return out


def _bookmarks(body):
    return _ranges(body, qn("w:bookmarkStart"), qn("w:bookmarkEnd"), lambda el: el.get(qn("w:name")))


def test_statuses(ann, run):
    _, reqs, _ = run
    statuses = {rid: [s.match.status for s in req.sources] for rid, req in reqs.items()}
    S = ann.Status
    assert statuses == {
        "REQ-001": [S.FOUND], "REQ-002": [S.FOUND, S.FOUND], "REQ-003": [S.FOUND, S.FOUND, S.FOUND],
        "REQ-004": [S.FOUND], "REQ-005": [S.NOT_FOUND], "REQ-006": [S.EXTERNAL],
    }


def test_bookmarks_cover_exactly_the_quotes(ann, run):
    out, _, _ = run
    bookmarks = _bookmarks(_body(out))
    assert set(bookmarks) == set(PASSAGES)
    for name, passage in PASSAGES.items():
        assert ann.compact(bookmarks[name]) == ann.compact(passage), name


def test_no_text_lost_or_added(word_case, run):
    out, _, _ = run
    assert _visible_text(_body(out), skip_hyperlinks=True) == _visible_text(_body(word_case[0]), skip_hyperlinks=False)


def test_tracked_change_kept(run):
    out, _, _ = run
    body = _body(out)
    assert [t.text for t in body.iter(qn("w:delText"))] == ["10"]


def test_only_the_passages_are_highlighted(ann, run):
    out, _, _ = run
    colours: dict[str, list[str]] = {}
    for r in _body(out).iter(qn("w:r")):
        h = r.find(f"{qn('w:rPr')}/{qn('w:highlight')}")
        if h is not None:
            colours.setdefault(h.get(qn("w:val")), []).append("".join(t.text or "" for t in r.iter(qn("w:t"))))
    yellow = "".join(PASSAGES[n] for n in ("REQ_001_1", "REQ_002_1", "REQ_003_1", "REQ_004_1", "REQ_003_2"))
    assert ann.compact("".join(colours["yellow"])) == ann.compact(yellow)
    assert ann.compact("".join(colours["cyan"])) == ann.compact(PASSAGES["REQ_002_2"])  # qualifier (turquoise)


def test_split_runs_keep_their_formatting(run):
    out, _, _ = run
    first_paragraph = next(p for p in docx.Document(out).paragraphs if p.text.startswith("Innledning"))
    bold = [r.text for r in first_paragraph.runs if r.bold]
    assert bold == ["14"]
    assert first_paragraph.runs[0].text == "Innledning. " and first_paragraph.runs[0].font.highlight_color is None


def test_markers_follow_the_passage_and_cycle(run):
    out, _, _ = run
    links = [("".join(t.text for t in h.iter(qn("w:t"))), h.get(qn("w:anchor")))
             for h in _body(out).iter(qn("w:hyperlink"))]
    assert links == [
        (" [REQ-001]", "REQ_001_1"),
        (" [REQ-002 1/2 »]", "REQ_002_2"),
        (" [REQ-003 1/3 »]", "REQ_003_2"),
        (" [REQ-004]", "REQ_004_1"),
        (" [REQ-003 2/3 »]", "REQ_003_3"),
        (" [REQ-002 2/2 »]", "REQ_002_1"),  # same passage: markers in requirement order
        (" [REQ-003 3/3 »]", "REQ_003_1"),
    ]


def test_comments(run):
    out, _, _ = run
    body = _body(out)
    comments = {c.comment_id: c.text for c in docx.Document(out).comments}
    ranges = _ranges(body, qn("w:commentRangeStart"), qn("w:commentRangeEnd"), lambda el: comments[int(el.get(qn("w:id")))])
    assert ranges["REQ-001"] == PASSAGES["REQ_001_1"]
    assert "REQ-002 (2/2), presisering – også: Kapittel 2 Purring, side 1" in ranges
    assert "REQ-003 (1/3) – også: Kapittel 3 Gebyrer, side 2, tabell 1, rad 3; Kapittel 3 Gebyrer, side 2" in ranges
    assert len(comments) == 7


def test_excel_links_point_to_bookmarks(run):
    out, reqs, xlsx = run
    ws = openpyxl.load_workbook(xlsx).active
    links = {rid: ws.cell(req.row, 3).hyperlink for rid, req in reqs.items()}
    assert (links["REQ-001"].target, links["REQ-001"].location) == (out.name, "REQ_001_1")
    assert links["REQ-003"].location == "REQ_003_1"
    assert links["REQ-005"] is None and links["REQ-006"] is None


def test_rerun_gives_the_same_result(ann, run, word_case):
    out, _, xlsx = run
    first = (_bookmarks(_body(out)), _marker_anchors(_body(out)))
    ann.annotate_files(word_case[0], xlsx)
    assert (_bookmarks(_body(out)), _marker_anchors(_body(out))) == first


def test_existing_bookmarks_keep_unique_ids(ann, word_case, tmp_path):
    source, xlsx = word_case
    d = docx.Document(source)
    p = d.paragraphs[0]._p
    start = p.makeelement(qn("w:bookmarkStart"), {qn("w:id"): "7", qn("w:name"): "Existing"})
    p.insert(0, start)
    p.append(p.makeelement(qn("w:bookmarkEnd"), {qn("w:id"): "7"}))
    d.save(source)
    out, _ = ann.annotate_files(source, xlsx)
    ids = [b.get(qn("w:id")) for b in _body(out).iter(qn("w:bookmarkStart"))]
    assert len(ids) == len(set(ids)) and "Existing" in _bookmarks(_body(out))


def test_main_prints_report(ann, word_case, capsys):
    ann.main([str(p) for p in word_case])
    out = capsys.readouterr().out
    assert "Sources: 9 | external: 1 | found: 7 | not found: 1" in out
    assert "REQ-005 source 1/1 [not found]" in out
