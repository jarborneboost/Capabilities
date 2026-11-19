"""The text matching core: normalising, patterns and span search."""
import pytest


@pytest.mark.parametrize("raw, expected", [
    ("“Quoted” – text", '"quoted"-text'),  # typographic quotes and dashes
    ("betalings­filer", "betalingsfiler"),  # soft hyphen
    ("• first item", "firstitem"),  # bullet
    ("Line\nbreak\ttab  spaces", "linebreaktabspaces"),
    ("wait…", "wait..."),
])
def test_compact_normalises(ann, raw, expected):
    assert ann.compact(raw) == expected


def test_pattern_for_text(ann):
    pat = ann.pattern_for("The system shall respond.")
    assert (pat.kind, pat.segments, pat.max_gap) == ("text", ["thesystemshallrespond."], 0)


def test_pattern_for_table_row_skips_empty_cells(ann):
    pat = ann.pattern_for("Purring | 14 |  | 100,00")
    assert pat.kind == "table"
    assert pat.segments == ["purring", "14", "100,00"]
    assert pat.max_gap == ann.TABLE_MAX_GAP


@pytest.mark.parametrize("quote", ["Start of it ... end of it", "Start of it … end of it",
                                   "Start of it [...] end of it"])
def test_pattern_for_ellipsis(ann, quote):
    pat = ann.pattern_for(quote)
    assert pat.kind == "ellipsis"
    assert pat.segments == ["startofit", "endofit"]


def test_find_spans_exact_and_several(ann):
    text = ann.compact("Alpha beta gamma. Alpha beta gamma.")
    pat = ann.pattern_for("alpha beta")
    assert ann.find_spans(text, pat) == [(0, 9), (15, 24)]


def test_find_spans_respects_gap(ann):
    pat = ann.pattern_for("Purring | 14")
    assert ann.find_spans("purring2x14", pat) == [(0, 11)]  # gap of 2 allowed
    assert ann.find_spans("purring" + "x" * (ann.TABLE_MAX_GAP + 1) + "14", pat) == []


def test_find_spans_stops_at_max_matches(ann):
    text = "ab" * 50
    assert len(ann.find_spans(text, ann.pattern_for("abab"))) == ann.MAX_MATCHES


def test_find_anchored_bridges_inserted_text(ann):
    quote = ann.compact("The customer portal shall show all open invoices, their due dates and the "
                        "amount that remains to be paid after partial payments.")
    assert len(quote) >= 3 * ann.ANCHOR_LEN
    broken = quote[:60] + "pagefooter28" + quote[60:]
    assert ann.find_spans(broken, ann.pattern_for(quote)) == []
    assert ann.find_anchored(broken, quote) == [(0, len(broken))]


def test_find_anchored_ignores_short_quotes(ann):
    short = "x" * (3 * ann.ANCHOR_LEN - 1)
    assert ann.find_anchored(short, short) == []


def test_search_statuses(ann):
    text = ann.compact("Some text in the document.")
    assert ann.search(text, "ab") is ann.Status.TOO_SHORT
    assert ann.search(text, "Other text entirely") is ann.Status.NOT_FOUND
    spans, status = ann.search(text, "text in the")
    assert status is ann.Status.FOUND and len(spans) == 1


@pytest.mark.parametrize("raw, expected", [
    ('"quoted"', "quoted"), ("“quoted”", "quoted"), ("«quoted»", "quoted"), ("unquoted", "unquoted"),
    ('"', '"'),
])
def test_strip_quote(ann, raw, expected):
    assert ann.strip_quote(raw) == expected


@pytest.mark.parametrize("rid, k, expected", [
    ("REQ-001", 1, "REQ_001_1"),
    ("REQ-PP1-001", 2, "REQ_PP1_001_2"),
    ("001", 1, "R_001_1"),
])
def test_bookmark_name(ann, rid, k, expected):
    assert ann.bookmark_name(rid, k) == expected


def test_bookmark_name_keeps_suffix_within_word_limit(ann):
    name = ann.bookmark_name("REQ-" + "X" * 60, 12)
    assert len(name) == ann.BOOKMARK_MAX_LEN
    assert name.endswith("_12")


def _req(ann, *sources):
    req = ann.Requirement("REQ-007", 2)
    for k, (marker, loc) in enumerate(sources, 1):
        req.sources.append(ann.Source("REQ-007", k, len(sources), marker, "q", loc))
    return req


def test_comment_text_single_source(ann):
    req = _req(ann, ("", "Section 3, page 8"))
    assert ann.comment_text(req.sources[0], req, ann.LANG_WORDS["en"]) == "REQ-007"


def test_comment_text_lists_other_sources(ann):
    req = _req(ann, ("*", "Section 3, page 8"), ("-", "Section 6, page 21"))
    assert ann.comment_text(req.sources[0], req, ann.LANG_WORDS["en"]) == "REQ-007 (1/2) – also: Section 6, page 21"
    assert ann.comment_text(req.sources[1], req, ann.LANG_WORDS["no"]) == \
        "REQ-007 (2/2), presisering – også: Section 3, page 8"


def test_next_found_cycles_over_found_sources(ann):
    req = _req(ann, ("*", "a"), ("*", "b"), ("-", "c"))
    a, b, c = req.sources
    a.match, c.match = ann.Match(ann.Status.FOUND), ann.Match(ann.Status.AMBIGUOUS)
    b.match = ann.Match(ann.Status.NOT_FOUND)
    assert req.next_found(a) is c and req.next_found(c) is a
    assert req.link_source() is a


def test_link_source_prefers_obligation(ann):
    req = _req(ann, ("*", "a"), ("-", "b"))
    req.sources[1].match = ann.Match(ann.Status.FOUND)
    assert req.link_source() is req.sources[1]  # the obligation wasn't found
    req.sources[0].match = ann.Match(ann.Status.FOUND)
    assert req.link_source() is req.sources[0]
