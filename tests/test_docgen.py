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
    CONFIGURATION_PLACEHOLDER,
    MAX_WEIGHT,
    MIN_WEIGHT,
    PLACEHOLDER,
    DocgenError,
    OutputLockedError,
    _section_index,
    build_document,
    column_weights,
    generate_declared_list,
)
from dlg.models import Item

EXAMPLE_TEMPLATE = Path(__file__).resolve().parent.parent / "templates" / "declared_list_template.docx"


COLUMNS = ["ID", "Name", "Description", "Category"]


def _item(item_id: str, name: str = "", description: str = "", category: str = "") -> Item:
    """Create an item with the four test columns."""
    values = {"ID": item_id, "Name": name, "Description": description, "Category": category}
    return Item(id=item_id, values=values)


def _items(n: int = 3) -> list[Item]:
    """Create ``n`` simple items."""
    return [_item(f"I{k}", f"Name {k}", f"Desc {k}", "Cat") for k in range(n)]


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
    doc = build_document(EXAMPLE_TEMPLATE, COLUMNS, _items(5))
    assert all(PLACEHOLDER not in p.text for p in doc.paragraphs)
    assert len(doc.tables) == 1
    table = doc.tables[0]
    assert len(table.rows) == 6  # header + 5
    section = doc.sections[_section_index(doc, table._tbl)]
    assert section.orientation == WD_ORIENT.LANDSCAPE
    assert _orientations(doc) == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE, WD_ORIENT.PORTRAIT]


def test_table_content_and_header() -> None:
    """Exactly the given columns, in order, with no extra numbering; header row repeats."""
    doc = build_document(EXAMPLE_TEMPLATE, COLUMNS, _items(2))
    table = doc.tables[0]
    assert [c.text for c in table.rows[0].cells] == COLUMNS
    assert [c.text for c in table.rows[2].cells] == ["I1", "Name 1", "Desc 1", "Cat"]
    assert table.rows[0]._tr.trPr.find(qn("w:tblHeader")) is not None
    assert table.rows[1]._tr.trPr.find(qn("w:tblHeader")) is None


def test_column_widths_fit_text_area() -> None:
    """Column widths sum to (about) the landscape text width."""
    doc = build_document(EXAMPLE_TEMPLATE, COLUMNS, _items(1))
    section = doc.sections[1]
    expected = section.page_width - section.left_margin - section.right_margin
    total = sum(c.width for c in doc.tables[0].rows[1].cells)
    assert abs(total - expected) < 10


def test_column_grid_matches_cell_widths() -> None:
    """The table grid (used by LibreOffice) agrees with the cell widths."""
    table = build_document(EXAMPLE_TEMPLATE, COLUMNS, _items(1)).tables[0]
    assert [c.width for c in table.columns] == [c.width for c in table.rows[1].cells]
    assert table._tbl.tblPr.find(qn("w:tblLayout")).get(qn("w:type")) == "fixed"


def test_item_order_is_preserved() -> None:
    """Rows follow the order of the items passed in."""
    items = list(reversed(_items(3)))
    table = build_document(EXAMPLE_TEMPLATE, COLUMNS, items).tables[0]
    assert [r.cells[0].text for r in table.rows[1:]] == ["I2", "I1", "I0"]


def test_special_and_control_characters() -> None:
    """XML-special characters survive; illegal control characters are stripped."""
    item = _item("X", "A & B <c>", "bad\x00char\x0b!")
    table = build_document(EXAMPLE_TEMPLATE, COLUMNS, [item]).tables[0]
    assert table.rows[1].cells[1].text == "A & B <c>"
    assert table.rows[1].cells[2].text == "badchar!"


def test_placeholder_split_across_runs(tmp_path: Path) -> None:
    """A placeholder Word has split into several runs is still found."""
    doc = Document()
    p = doc.add_paragraph()
    for part in ("{{DECLARED", "_LIST", "}}"):
        p.add_run(part)
    path = tmp_path / "t.docx"
    doc.save(path)
    assert len(build_document(path, COLUMNS, _items(1)).tables) == 1


def test_placeholder_holding_section_break_keeps_sections(tmp_path: Path) -> None:
    """If the placeholder paragraph carries the section break, it is not deleted."""
    template = _make_template(tmp_path / "t.docx", placeholder_carries_break=True)
    doc = build_document(template, COLUMNS, _items(2))
    assert _orientations(doc) == [WD_ORIENT.PORTRAIT, WD_ORIENT.LANDSCAPE, WD_ORIENT.PORTRAIT]
    assert all(PLACEHOLDER not in p.text for p in doc.paragraphs)
    section = doc.sections[_section_index(doc, doc.tables[0]._tbl)]
    assert section.orientation == WD_ORIENT.LANDSCAPE


def test_errors(tmp_path: Path) -> None:
    """Empty selection, missing/invalid template and missing placeholder all raise."""
    with pytest.raises(DocgenError, match="No items"):
        build_document(EXAMPLE_TEMPLATE, COLUMNS, [])
    with pytest.raises(DocgenError, match="no columns"):
        build_document(EXAMPLE_TEMPLATE, [], _items())
    with pytest.raises(DocgenError, match="not found"):
        build_document(tmp_path / "nope.docx", COLUMNS, _items())
    bad = tmp_path / "bad.docx"
    bad.write_text("not a docx")
    with pytest.raises(DocgenError, match="Cannot read"):
        build_document(bad, COLUMNS, _items())
    plain = tmp_path / "plain.docx"
    Document().save(plain)
    with pytest.raises(DocgenError, match="no paragraph"):
        build_document(plain, COLUMNS, _items())


def test_generate_writes_file_and_creates_folders(tmp_path: Path) -> None:
    """The document is saved (creating folders) and no temp file remains."""
    out = generate_declared_list(EXAMPLE_TEMPLATE, COLUMNS, _items(2), tmp_path / "sub" / "out.docx")
    assert len(Document(str(out)).tables[0].rows) == 3
    assert [p.name for p in (tmp_path / "sub").iterdir()] == ["out.docx"]


def test_locked_output_file(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """A file that cannot be replaced (open in Word on Windows) gives a clear error; temp file cleaned up."""
    import dlg.docgen as docgen

    def deny(src: object, dst: object) -> None:
        raise PermissionError(13, "Permission denied")

    monkeypatch.setattr(docgen.os, "replace", deny)
    with pytest.raises(OutputLockedError, match="open in Word"):
        generate_declared_list(EXAMPLE_TEMPLATE, COLUMNS, _items(1), tmp_path / "out.docx")
    assert list(tmp_path.iterdir()) == []


def _template_with_config_placeholders(path: Path) -> Path:
    """A template using {{CONFIGURATION_NAME}} in body, split runs, a table and the header."""
    doc = Document()
    doc.add_paragraph(f"Config: {CONFIGURATION_PLACEHOLDER}.")
    split = doc.add_paragraph()
    for part in ("For {{CONFIG", "URATION_", "NAME}} only"):
        split.add_run(part)
    split.runs[0].bold = True
    doc.add_table(rows=1, cols=1).rows[0].cells[0].text = f"Cell {CONFIGURATION_PLACEHOLDER}"
    doc.sections[0].header.paragraphs[0].text = f"Header {CONFIGURATION_PLACEHOLDER}"
    doc.add_paragraph(PLACEHOLDER)
    doc.save(path)
    return path


def test_configuration_name_replaced_everywhere(tmp_path: Path) -> None:
    """Body, split runs, table cells and headers are filled; item text is never substituted."""
    template = _template_with_config_placeholders(tmp_path / "t.docx")
    item = _item("X", f"literal {CONFIGURATION_PLACEHOLDER}")
    doc = build_document(template, COLUMNS, [item], "Rig <7> & B")
    texts = [p.text for p in doc.paragraphs]
    assert "Config: Rig <7> & B." in texts
    assert "For Rig <7> & B only" in texts
    assert doc.tables[0].rows[0].cells[0].text == "Cell Rig <7> & B"
    assert doc.sections[0].header.paragraphs[0].text == "Header Rig <7> & B"
    item_table = doc.tables[1]
    assert item_table.rows[1].cells[1].text == f"literal {CONFIGURATION_PLACEHOLDER}"


def test_configuration_placeholder_in_single_run_keeps_formatting(tmp_path: Path) -> None:
    """A placeholder inside one run is replaced in that run, keeping its formatting."""
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Name: ")
    bold = p.add_run(CONFIGURATION_PLACEHOLDER)
    bold.bold = True
    doc.add_paragraph(PLACEHOLDER)
    doc.save(tmp_path / "t.docx")
    out = build_document(tmp_path / "t.docx", COLUMNS, _items(1), "Rig")
    runs = out.paragraphs[0].runs
    assert [(r.text, bool(r.bold)) for r in runs] == [("Name: ", False), ("Rig", True)]


def test_empty_configuration_name_and_linked_headers(tmp_path: Path) -> None:
    """No name leaves an empty value; linked headers are not given content of their own."""
    template = _template_with_config_placeholders(tmp_path / "t.docx")
    doc = build_document(template, COLUMNS, _items(1))
    assert "Config: ." in [p.text for p in doc.paragraphs]
    plain = build_document(EXAMPLE_TEMPLATE, COLUMNS, _items(1), "Rig")
    assert [s.header.is_linked_to_previous for s in plain.sections] == [True, True, True]


def test_template_file_not_modified(tmp_path: Path) -> None:
    """Generating a document never alters the template on disk."""
    before = EXAMPLE_TEMPLATE.read_bytes()
    generate_declared_list(EXAMPLE_TEMPLATE, COLUMNS, _items(), tmp_path / "o.docx")
    assert EXAMPLE_TEMPLATE.read_bytes() == before


def test_table_properties_follow_schema_order() -> None:
    """``w:tblPr`` children are in OOXML schema order (Word is strict about it)."""
    order = ["tblStyle", "tblW", "jc", "tblBorders", "shd", "tblLayout", "tblLook"]
    tbl_pr = build_document(EXAMPLE_TEMPLATE, COLUMNS, _items(1)).tables[0]._tbl.tblPr
    names = [child.tag.split("}")[1] for child in tbl_pr]
    positions = [order.index(n) for n in names]
    assert positions == sorted(positions), names
    assert "tblBorders" in names and "tblW" in names


def test_any_columns_and_missing_values() -> None:
    """Arbitrary column names work; values missing from an item become empty cells."""
    items = [Item(id="1", values={"Part": "1", "Serial": "S-9"}), Item(id="2", values={"Part": "2"})]
    table = build_document(EXAMPLE_TEMPLATE, ["Part", "Serial", "Mass"], items).tables[0]
    assert [c.text for c in table.rows[0].cells] == ["Part", "Serial", "Mass"]
    assert [c.text for c in table.rows[2].cells] == ["2", "", ""]


def test_column_weights_follow_content() -> None:
    """Long text gets a wide column, short codes a narrow one, both within limits."""
    items = [_item("A-1", "Pump", "x " * 200, "C"), _item("A-22", "Valve", "short", "C")]
    w_id, w_name, w_desc, w_cat = column_weights(COLUMNS, items)
    assert w_desc == MAX_WEIGHT
    assert w_cat == len("Category")  # the header's longest word wins over the 1-char values
    assert MIN_WEIGHT <= w_id < w_name < w_desc
    assert column_weights(["X"], [Item(id="1", values={})]) == [MIN_WEIGHT]
