"""Reading the requirements sheet and writing the Location links."""
import openpyxl
import pytest


def _read(ann, make_xlsx, tmp_path, rows, lang="en"):
    ws = openpyxl.load_workbook(make_xlsx(tmp_path / "r.xlsx", rows, lang)).active
    return ann.read_requirements(ws)


def test_reads_markers_and_locations(ann, make_xlsx, tmp_path):
    cols, lang, reqs = _read(ann, make_xlsx, tmp_path, [(
        "REQ-001",
        '* "The obligation."\n- "The definition."',
        "* Section 3 Search, page 8\n- Section 6 Definitions, page 21, table 4, row 3",
    )])
    assert (cols, lang) == ({"id": 1, "text": 2, "loc": 3}, "en")
    first, second = reqs[0].sources
    assert (first.marker, first.quote, first.location) == ("*", "The obligation.", "Section 3 Search, page 8")
    assert (second.k, second.n, second.marker) == (2, 2, "-")
    assert second.location == "Section 6 Definitions, page 21, table 4, row 3"
    assert second.page_hint() == [20]
    assert reqs[0].row == 2 and first.match is None and first.warning == ""


def test_single_source_without_marker(ann, make_xlsx, tmp_path):
    _, _, reqs = _read(ann, make_xlsx, tmp_path, [("REQ-001", '"Just one."', "Section 1, page 2")])
    src = reqs[0].sources[0]
    assert (src.marker, src.quote, src.n) == ("", "Just one.", 1)


def test_page_range_hint(ann, make_xlsx, tmp_path):
    _, _, reqs = _read(ann, make_xlsx, tmp_path, [("REQ-001", '"x y z w"', "Section 5.3, page 28-29")])
    assert reqs[0].sources[0].page_hint() == [27, 28]


def test_count_mismatch_gives_warning(ann, make_xlsx, tmp_path):
    _, _, reqs = _read(ann, make_xlsx, tmp_path, [("REQ-001", '* "One."\n* "Two."', "Section 1, page 1")])
    assert all(s.warning == "2 quotes but 1 locations" for s in reqs[0].sources)
    assert reqs[0].sources[1].location == ""


def test_norwegian_headers(ann, make_xlsx, tmp_path):
    _, lang, reqs = _read(ann, make_xlsx, tmp_path, [("REQ-001", '"Tekst her."', "Kapittel 1, side 3")], lang="no")
    assert lang == "no"
    assert reqs[0].sources[0].page_hint() == [2]


def test_url_location_is_external(ann, make_xlsx, tmp_path):
    _, _, reqs = _read(ann, make_xlsx, tmp_path, [("REQ-001", '"From the website."', "https://example.com/a")])
    assert reqs[0].sources[0].match.status is ann.Status.EXTERNAL


def test_rows_without_id_are_skipped(ann, make_xlsx, tmp_path):
    _, _, reqs = _read(ann, make_xlsx, tmp_path, [("REQ-001", '"A."', "p"), ("", "", ""), ("REQ-002", '"B."', "p")])
    assert [r.rid for r in reqs] == ["REQ-001", "REQ-002"]


def test_missing_column_raises(ann, tmp_path):
    wb = openpyxl.Workbook()
    wb.active.append(["ID", "Original text"])
    with pytest.raises(ann.AnnotateError, match="Column not found"):
        ann.read_requirements(wb.active)


def test_write_links_sets_and_clears(ann, make_xlsx, tmp_path):
    path = make_xlsx(tmp_path / "r.xlsx", [("REQ-001", '"A."', "page 1"), ("REQ-002", '"B."', "page 2")])
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    cols, _, reqs = ann.read_requirements(ws)
    for src in (reqs[0].sources[0], reqs[1].sources[0]):
        src.match = ann.Match(ann.Status.FOUND)
    reqs[0].sources[0].placed = ann.Placed(link="REQ_001_1")
    reqs[1].sources[0].placed = ann.Placed(link="REQ_002_1")
    ann.write_links(ws, cols["loc"], reqs, "doc - Annotated.docx")
    cell = ws.cell(2, cols["loc"])
    assert (cell.hyperlink.target, cell.hyperlink.location) == ("doc - Annotated.docx", "REQ_001_1")
    assert cell.font.underline == "single"

    # a rerun where REQ-002 is no longer found: its old link and link styling go away
    reqs[1].sources[0].match, reqs[1].sources[0].placed = ann.Match(ann.Status.NOT_FOUND), None
    ann.write_links(ws, cols["loc"], reqs, "doc - Annotated.docx")
    stale = ws.cell(3, cols["loc"])
    assert stale.hyperlink is None
    assert stale.font.underline is None
