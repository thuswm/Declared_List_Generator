"""Tests for Word document generation."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.document import Document as DocumentType
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

from dlg.docgen import (
    PLACEHOLDER,
    DocgenError,
    _section_index,
    build_document,
    generate_declared_list,
    render_to_bytes,
)
from dlg.models import Item

EXAMPLE_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "declared_list_template.docx"


def _items(n: int = 3) -> list[Item]:
    """Create ``n`` simple items."""
    return [Item(id=f"I{k}", name=f"Name {k}", description=f"Desc {k}", category="Cat") for k in range(n)]


def _orientations(doc: DocumentType) -> list[WD_ORIENT]:
    """Return the orientation of each section of a document."""
    return [s.orientation for s in doc.sections]


def _make_template(path: Path, placeholder_carries_break: bool = False) -> Path:
    """Build a portrait / landscape / portrait template in code.

    Args:
        path: Where to save it.
        placeholder_carries_break: If True, the placeholder is the last
            paragraph of the landscape section (so it holds the section break),
            as Word produces when the break directly follows the placeholder.
    """
    doc = Document()
    doc.add_paragraph("Intro")
    land = doc.add_section(WD_SECTION.NEW_PAGE)
    land.orientation = WD_ORIENT.LANDSCAPE
    land.page_width, land.page_height = land.page_height, land.page_width
    if not placeholder_carries_break:
        doc.add_paragraph(PLACEHOLDER)
    back = doc.add_section(WD_SECTION.NEW_PAGE)
    back.orientation = WD_ORIENT.PORTRAIT
    back.page_width, back.page_height = back.page_height, back.page_width
    doc.add_paragraph("Outro")
    if placeholder_carries_break:
        # The empty paragraph ending the landscape section holds its sectPr.
        for p in doc.paragraphs:
            if p._p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None and not p.text:
                if _is_landscape_break(p):
                    p.add_run(PLACEHOLDER)
                    break
    doc.save(path)
    return path


def _is_landscape_break(paragraph: Paragraph) -> bool:
    """Return True if a paragraph's section break describes a landscape page."""
    sect = paragraph._p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
    pg = sect.find(qn("w:pgSz"))
    return int(pg.get(qn("w:w"))) > int(pg.get(qn("w:h")))


def test_example_template_layout() -> None:
    """The example template is portrait, landscape, portrait."""
    doc = Document(str(EXAMPLE_TEMPLATE))
    assert _orientations(doc) == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE, WD_ORIENT.PORTRAIT]


def test_table_replaces_placeholder_in_landscape_section() -> None:
    """The table lands in the landscape section; orientations are unchanged."""
    doc = build_document(EXAMPLE_TEMPLATE, _items(5))
    assert all(PLACEHOLDER not in p.text for p in doc.paragraphs)
    assert len(doc.tables) == 1
    table = doc.tables[0]
    assert len(table.rows) == 6  # header + 5
    section = doc.sections[_section_index(doc, table._tbl)]
    assert section.orientation == WD_ORIENT.LANDSCAPE
    assert _orientations(doc) == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE, WD_ORIENT.PORTRAIT]


def test_table_content_and_header() -> None:
    """Headers, numbering and cell text are correct; header row repeats."""
    doc = build_document(EXAMPLE_TEMPLATE, _items(2))
    table = doc.tables[0]
    assert [c.text for c in table.rows[0].cells] == ["No.", "ID", "Name", "Description", "Category"]
    assert [c.text for c in table.rows[2].cells] == ["2", "I1", "Name 1", "Desc 1", "Cat"]
    assert table.rows[0]._tr.trPr.find(qn("w:tblHeader")) is not None
    assert table.rows[1]._tr.trPr.find(qn("w:tblHeader")) is None


def test_column_widths_fit_text_area() -> None:
    """Column widths sum to (about) the landscape text width."""
    doc = build_document(EXAMPLE_TEMPLATE, _items(1))
    section = doc.sections[1]
    expected = section.page_width - section.left_margin - section.right_margin
    total = sum(c.width for c in doc.tables[0].rows[1].cells)
    assert abs(total - expected) < 10


def test_column_grid_matches_cell_widths() -> None:
    """The table grid (used by LibreOffice) agrees with the cell widths."""
    table = build_document(EXAMPLE_TEMPLATE, _items(1)).tables[0]
    assert [c.width for c in table.columns] == [c.width for c in table.rows[1].cells]
    assert table._tbl.tblPr.find(qn("w:tblLayout")).get(qn("w:type")) == "fixed"


def test_item_order_is_preserved() -> None:
    """Rows follow the order of the items passed in."""
    items = list(reversed(_items(3)))
    table = build_document(EXAMPLE_TEMPLATE, items).tables[0]
    assert [r.cells[1].text for r in table.rows[1:]] == ["I2", "I1", "I0"]


def test_special_and_control_characters() -> None:
    """XML-special characters survive; illegal control characters are stripped."""
    item = Item(id="X", name="A & B <c>", description="bad\x00char\x0b!", category="")
    table = build_document(EXAMPLE_TEMPLATE, [item]).tables[0]
    assert table.rows[1].cells[2].text == "A & B <c>"
    assert table.rows[1].cells[3].text == "badchar!"


def test_placeholder_split_across_runs(tmp_path: Path) -> None:
    """A placeholder Word has split into several runs is still found."""
    doc = Document()
    p = doc.add_paragraph()
    for part in ("{{DECLARED", "_LIST", "}}"):
        p.add_run(part)
    path = tmp_path / "t.docx"
    doc.save(path)
    assert len(build_document(path, _items(1)).tables) == 1


def test_placeholder_holding_section_break_keeps_sections(tmp_path: Path) -> None:
    """If the placeholder paragraph carries the section break, it is not deleted."""
    template = _make_template(tmp_path / "t.docx", placeholder_carries_break=True)
    doc = build_document(template, _items(2))
    assert _orientations(doc) == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE, WD_ORIENT.PORTRAIT]
    assert all(PLACEHOLDER not in p.text for p in doc.paragraphs)
    section = doc.sections[_section_index(doc, doc.tables[0]._tbl)]
    assert section.orientation == WD_ORIENT.LANDSCAPE


def test_errors(tmp_path: Path) -> None:
    """Empty selection, missing/invalid template and missing placeholder all raise."""
    with pytest.raises(DocgenError, match="No items"):
        build_document(EXAMPLE_TEMPLATE, [])
    with pytest.raises(DocgenError, match="not found"):
        build_document(tmp_path / "nope.docx", _items())
    bad = tmp_path / "bad.docx"
    bad.write_text("not a docx")
    with pytest.raises(DocgenError, match="Cannot read"):
        build_document(bad, _items())
    plain = tmp_path / "plain.docx"
    Document().save(plain)
    with pytest.raises(DocgenError, match="no paragraph"):
        build_document(plain, _items())


def test_generate_and_render_to_bytes(tmp_path: Path) -> None:
    """Saving to disk and to memory both yield a readable document."""
    out = generate_declared_list(EXAMPLE_TEMPLATE, _items(2), tmp_path / "sub" / "out.docx")
    assert len(Document(str(out)).tables[0].rows) == 3
    data = render_to_bytes(EXAMPLE_TEMPLATE, _items(2))
    assert data[:2] == b"PK"  # a .docx is a zip file


def test_template_file_not_modified(tmp_path: Path) -> None:
    """Generating a document never alters the template on disk."""
    before = EXAMPLE_TEMPLATE.read_bytes()
    generate_declared_list(EXAMPLE_TEMPLATE, _items(), tmp_path / "o.docx")
    assert EXAMPLE_TEMPLATE.read_bytes() == before


def test_table_properties_follow_schema_order() -> None:
    """``w:tblPr`` children are in OOXML schema order (Word is strict about it)."""
    order = ["tblStyle", "tblW", "jc", "tblBorders", "shd", "tblLayout", "tblLook"]
    tbl_pr = build_document(EXAMPLE_TEMPLATE, _items(1)).tables[0]._tbl.tblPr
    names = [child.tag.split("}")[1] for child in tbl_pr]
    positions = [order.index(n) for n in names]
    assert positions == sorted(positions), names
    assert "tblBorders" in names and "tblW" in names
