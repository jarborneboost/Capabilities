"""Fixtures for the extract-requirements annotation script.

All documents are generated here, with invented text that copies the patterns seen in real
customer documents (runs split mid-word, tracked changes, tables, review comments in the
margin of a PDF, passages crossing a page break). No customer documents in the repo.
"""
import importlib.util
import sys
from pathlib import Path

import docx
import openpyxl
import pymupdf
import pytest
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SCRIPT = (Path(__file__).parents[3] / "plugins" / "heistwright" / "skills" / "extract-requirements"
          / "scripts" / "annotate.py")

HEADERS = {
    "en": ["ID", "Original text", "Location", "ISTQB issues", "Rewritten requirement", "Testability",
           "Testability reason"],
    "no": ["ID", "Opprinnelig tekst", "Plassering", "ISTQB-avvik", "Omskrevet krav", "Testbarhet",
           "Begrunnelse for testbarhet"],
}


@pytest.fixture(scope="session")
def ann():
    """The annotate.py script, imported as a module."""
    spec = importlib.util.spec_from_file_location("annotate", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    sys.modules["annotate"] = module  # dataclasses look up their module while being created
    spec.loader.exec_module(module)
    return module


def write_xlsx(path: Path, rows: list[tuple[str, str, str]], lang: str = "en") -> Path:
    """Requirements sheet with (ID, Original text, Location) rows; other columns filled in."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(HEADERS[lang])
    for rid, text, loc in rows:
        ws.append([rid, text, loc, "None", "x", "OK", ""])
    wb.save(path)
    return path


@pytest.fixture
def make_xlsx():
    return write_xlsx


# ----------------------------------------------------------------- PDF

PDF_WIDTH, PDF_HEIGHT = 612, 792
BODY_X0, BODY_X1, MARGIN_X0 = 72, 400, 440

PDF_TEXT = {
    "invoice": "The system shall send every invoice through the print channel.",
    "duplicate": "Payment reminders are sent by email.",
    "comment": "Comment [AB1]: Reminders must also be sent by post.",
    "break_start": "The customer portal shall show all open invoices, their due dates and the",
    "break_end": "amount that remains to be paid after partial payments have been registered.",
    "page2_comment": "Comment [CD2]: Should partial payments be shown per invoice or in total?",
}
PDF_TABLE = [("Type", "Days", "Fee"), ("Reminder", "14", "38,00"), ("Notice", "33", "38,00")]


def _body(page, y: float, text: str, size: float = 11) -> float:
    rect = pymupdf.Rect(BODY_X0, y, BODY_X1, y + 60)
    assert page.insert_textbox(rect, text, fontsize=size) >= 0
    lines = len(page.get_textbox(rect).splitlines()) or 1
    return y + lines * size * 1.3 + 12


def _margin(page, y: float, text: str) -> None:
    rect = pymupdf.Rect(MARGIN_X0, y, PDF_WIDTH - 20, y + 60)
    page.draw_rect(rect, color=(0.8, 0.8, 0.8), fill=(0.95, 0.95, 0.95))
    assert page.insert_textbox(rect + (3, 3, -3, -3), text, fontsize=8) >= 0


def _page_number(page, n: int) -> None:
    # like a Word export: the page number sits bottom right, right of the body text
    page.insert_text((540, PDF_HEIGHT - 40), str(n), fontsize=10)


@pytest.fixture
def pdf_doc(tmp_path) -> Path:
    """3 pages: body text with a review-comment column on the right (pages 1-2), a passage
    crossing the page break, a table, and one sentence that appears twice (pages 1 and 3)."""
    doc = pymupdf.open()
    p1 = doc.new_page(width=PDF_WIDTH, height=PDF_HEIGHT)
    y = _body(p1, 72, "1 Invoicing", 14)
    _margin(p1, y, PDF_TEXT["comment"])
    y = _body(p1, y, PDF_TEXT["invoice"])
    y = _body(p1, y, PDF_TEXT["duplicate"])
    _body(p1, PDF_HEIGHT - 120, PDF_TEXT["break_start"])
    _page_number(p1, 1)

    p2 = doc.new_page(width=PDF_WIDTH, height=PDF_HEIGHT)
    _margin(p2, 72, PDF_TEXT["page2_comment"])
    y = _body(p2, 72, PDF_TEXT["break_end"])
    y = _body(p2, y + 10, "2 Fees", 14)
    for row in PDF_TABLE:  # one text object per cell, row by row, like a real table
        for x, cell in zip((BODY_X0, BODY_X0 + 120, BODY_X0 + 200), row):
            p2.insert_text((x, y), cell, fontsize=10)
        y += 16
    _page_number(p2, 2)

    p3 = doc.new_page(width=PDF_WIDTH, height=PDF_HEIGHT)
    _body(p3, 72, "3 Reminders", 14)
    _body(p3, 110, PDF_TEXT["duplicate"])
    _page_number(p3, 3)

    path = tmp_path / "Customer document.pdf"
    doc.save(path)
    return path


PDF_ROWS = [
    ("REQ-001", f'"{PDF_TEXT["invoice"]}"', "Section 1 Invoicing, page 1"),
    ("REQ-002",
     f'* "{PDF_TEXT["break_start"]} {PDF_TEXT["break_end"]}"\n- "{PDF_TEXT["page2_comment"]}"',
     "* Section 1 Invoicing, page 1-2\n- Section 1 Invoicing, page 2"),
    ("REQ-003", '"Reminder | 14 | 38,00"', "Section 2 Fees, page 2, table 1, row 2"),
    ("REQ-004", '"Reminders must also be sent by post."', "Section 1 Invoicing, page 1"),
    ("REQ-005", f'"{PDF_TEXT["duplicate"]}"', "Section 3 Reminders, page 3"),
    ("REQ-006", '"The system shall fly to the moon."', "Section 1 Invoicing, page 1"),
    ("REQ-007", '"14"', "Section 2 Fees, page 2"),
    ("REQ-008", f'"{PDF_TEXT["duplicate"]}"', "Section 3 Reminders"),
]


@pytest.fixture
def pdf_case(tmp_path, pdf_doc) -> tuple[Path, Path]:
    """The PDF fixture and an English requirements sheet citing it: a plain quote (REQ-001),
    a passage crossing the page break plus a margin comment (REQ-002), a table row (REQ-003),
    part of a margin comment (REQ-004), a sentence on pages 1 and 3 cited with page 3
    (REQ-005) and without a page (REQ-008), a wrong quote (REQ-006) and a quote too short
    to search for (REQ-007)."""
    return pdf_doc, write_xlsx(tmp_path / "Customer document - Requirements.xlsx", PDF_ROWS)


# ----------------------------------------------------------------- Word

WORD_FILLER = "Filler text to give the document some length. " * 8


def _tracked_change(par, deleted: str, inserted: str) -> None:
    d = OxmlElement("w:del")
    d.set(qn("w:id"), "901")
    d.set(qn("w:author"), "Customer")
    r, t = OxmlElement("w:r"), OxmlElement("w:delText")
    t.text = deleted
    r.append(t)
    d.append(r)
    par._p.append(d)
    i = OxmlElement("w:ins")
    i.set(qn("w:id"), "902")
    i.set(qn("w:author"), "Customer")
    r, t = OxmlElement("w:r"), OxmlElement("w:t")
    t.text = inserted
    r.append(t)
    i.append(r)
    par._p.append(i)


@pytest.fixture
def word_doc(tmp_path) -> Path:
    """A sentence split mid-word across runs with bold in the middle, a bulleted list, a
    tracked change (10 -> 14), a page break, a table and a sentence cited twice."""
    d = docx.Document()
    d.add_heading("1 Fakturering", 1)
    p = d.add_paragraph()
    for text, bold in (("Innledning. Fakturaen skal sen", False), ("des innen ", False), ("14", True),
                       (" dager etter forfall. Dette er etterfølgende tekst.", False)):
        p.add_run(text).bold = bold
    d.add_paragraph(WORD_FILLER)
    d.add_heading("2 Purring", 1)
    d.add_paragraph("Purringer skal inneholde følgende:")
    for item in ("saksnummer", "forfallsdato", "beløp"):
        d.add_paragraph(item, style="List Bullet")
    p = d.add_paragraph()
    p.add_run("Inkassovarsel skal sendes etter ")
    _tracked_change(p, "10", "14")
    p.add_run(" dager.")
    d.add_page_break()
    d.add_heading("3 Gebyrer", 1)
    table = d.add_table(rows=3, cols=4)
    for i, row in enumerate([("Type", "Dager", "Gebyr", "Grense"), ("Purring", "14", "38,00", ""),
                             ("Inkassovarsel", "33", "38,00", "100,00")]):
        for j, value in enumerate(row):
            table.cell(i, j).text = value
    d.add_paragraph("Gebyret skal være 38 kroner for både purring og inkassovarsel.")
    d.add_paragraph("Denne setningen finnes ikke i kravlisten.")
    path = tmp_path / "Kundedokument.docx"
    d.save(path)
    return path


WORD_ROWS = [
    ("REQ-001", '"Fakturaen skal sendes innen 14 dager etter forfall."', "Kapittel 1 Fakturering, side 1"),
    ("REQ-002",
     '* "Purringer skal inneholde følgende: saksnummer forfallsdato beløp"\n'
     '- "Gebyret skal være 38 kroner for både purring og inkassovarsel."',
     "* Kapittel 2 Purring, side 1\n- Kapittel 3 Gebyrer, side 2"),
    ("REQ-003",
     '* "Inkassovarsel skal sendes etter 14 dager."\n'
     '* "Inkassovarsel | 33 | 38,00 | 100,00"\n'
     '- "Gebyret skal være 38 kroner for både purring og inkassovarsel."',
     "* Kapittel 2 Purring, side 1\n* Kapittel 3 Gebyrer, side 2, tabell 1, rad 3\n- Kapittel 3 Gebyrer, side 2"),
    ("REQ-004", '"Purring | 14 | 38,00 | "', "Kapittel 3 Gebyrer, side 2, tabell 1, rad 2"),
    ("REQ-005", '"Fakturaen skal sendes innen 30 dager."', "Kapittel 1 Fakturering, side 1"),
    ("REQ-006", '"Se vedlegget på nettsiden."', "https://example.com/vedlegg"),
]


@pytest.fixture
def word_case(tmp_path, word_doc) -> tuple[Path, Path]:
    """The Word fixture and a Norwegian requirements sheet citing it."""
    return word_doc, write_xlsx(tmp_path / "Kundedokument - Requirements.xlsx", WORD_ROWS, lang="no")
