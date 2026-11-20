"""Tests for the Word table layout sheet (``dlg.layout``) and its use in docgen and the API."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from fastapi.testclient import TestClient
from openpyxl import Workbook

from dlg.app import create_app
from dlg.config import Config
from dlg.docgen import build_document
from dlg.layout import LayoutError, apply_layout, load_layout
from dlg.models import Column, Item, ItemTable
from dlg.sources.excel import ExcelItemSource

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "declared_list_template.docx"
LOCAL = ("127.0.0.1", 50000)

DATA = [
    ["ID", "Name", "Volt", "Amp", "#Secret"],
    ["P-1", "Pump", "230", "5", "x"],
    ["P-2", "Fan", "24", "1", "y"],
]


def layout_rows(ws, rows, merges=(), widths=None) -> None:
    """Fill a layout sheet: rows of cell values, merged ranges and column widths."""
    for row in rows:
        ws.append(row)
    for rng in merges:
        ws.merge_cells(rng)
    for letter, width in (widths or {}).items():
        ws.column_dimensions[letter].width = width


def make_workbook(path: Path, rows=None, merges=(), widths=None, layout_first: bool = False,
                  sheet: str = "Word layout") -> Path:
    """Write a database workbook with a data sheet and (optionally) a layout sheet."""
    wb = Workbook()
    data = wb.active
    data.title = "Items"
    for row in DATA:
        data.append(row)
    if rows is not None:
        ws = wb.create_sheet(sheet, 0 if layout_first else None)
        layout_rows(ws, rows, merges, widths)
    wb.save(path)
    return path


# The sketch from the user's notes: a group over two columns, one column title spanning all levels.
SKETCH = [
    ["Electrical", None, "Doc ref"],
    ["Voltage", "Current", None],
    ["Volt", "Amp", "Ref*"],
]
SKETCH_MERGES = ("A1:B1", "C1:C2")


def test_parse_multi_level_header(tmp_path: Path) -> None:
    """Header rows, merged cells, mapping row (hidden) and widths are read."""
    path = make_workbook(tmp_path / "db.xlsx", SKETCH, SKETCH_MERGES, {"A": 20, "B": 10})
    layout = load_layout(path, "Word layout", required=False)
    assert [(c.name, c.user_input) for c in layout.columns] == [("Volt", False), ("Amp", False), ("Ref", True)]
    assert layout.header_rows == 2
    cells = [(c.row, c.col, c.text, c.row_span, c.col_span) for c in layout.cells]
    assert cells == [
        (0, 0, "Electrical", 1, 2),
        (0, 2, "Doc ref", 2, 1),
        (1, 0, "Voltage", 1, 1),
        (1, 1, "Current", 1, 1),
    ]
    assert layout.widths == [20.0, 10.0, 8.43]  # C has no width: Excel's default
    assert layout.source == "db.xlsx › Word layout"


def test_only_mapping_row_gives_one_row_header(tmp_path: Path) -> None:
    """Without rows above, the mapping names are the header."""
    path = make_workbook(tmp_path / "db.xlsx", [["Name", "Volt"]])
    layout = load_layout(path, "word LAYOUT", required=False)  # name matched ignoring case
    assert layout.header_rows == 1
    assert [(c.text, c.col) for c in layout.cells] == [("Name", 0), ("Volt", 1)]


def test_table_starts_at_first_mapping_column_and_ignores_far_merges(tmp_path: Path) -> None:
    """Leading empty columns and unrelated merged notes outside the table are ignored."""
    rows = [
        [None, "Group", None, None, None, "note"],
        [None, "Volt", "Amp"],
    ]
    path = make_workbook(tmp_path / "db.xlsx", rows, ("B1:C1", "F1:G1"))
    layout = load_layout(path, "Word layout", required=False)
    assert layout.column_names == ["Volt", "Amp"]
    assert [(c.text, c.col_span) for c in layout.cells] == [("Group", 2)]


def test_missing_sheet(tmp_path: Path) -> None:
    """No sheet: automatic table, unless a separate layout file was set (then an error)."""
    path = make_workbook(tmp_path / "db.xlsx")
    assert load_layout(path, "Word layout", required=False) is None
    with pytest.raises(LayoutError, match="Sheets: Items"):
        load_layout(path, "Word layout", required=True)


@pytest.mark.parametrize("rows, merges, fragment", [
    ([["Volt", "Volt"]], (), "mapped twice"),
    ([["X"], ["*"]], (), "no column name"),
    ([["-*"]], (), "cannot be an input column"),
    ([["Group", None], ["Volt", "Amp"]], ("A1:A2",), "include the mapping row"),
    ([["Group", None, None], ["Volt", "Amp"]], ("A1:C1",), "outside the table"),
])
def test_layout_rule_errors(tmp_path: Path, rows, merges, fragment: str) -> None:
    """Broken layouts give an error that says what to fix."""
    path = make_workbook(tmp_path / "db.xlsx", rows, merges)
    with pytest.raises(LayoutError, match=fragment):
        load_layout(path, "Word layout", required=False)


def test_empty_layout_sheet_is_an_error(tmp_path: Path) -> None:
    """An existing but empty layout sheet is reported, not silently ignored."""
    path = make_workbook(tmp_path / "db.xlsx", [])
    with pytest.raises(LayoutError, match="empty"):
        load_layout(path, "Word layout", required=False)


def _table() -> ItemTable:
    return ItemTable(
        id_column="ID",
        columns=[Column(name="ID"), Column(name="Name"), Column(name="Volt")],
        items=[Item(id="P-1", values={"ID": "P-1", "Name": "Pump", "Volt": "230"})],
    )


def test_apply_layout_adds_and_marks_input_columns(tmp_path: Path) -> None:
    """Mapped '*' columns become input columns; missing ones are added; others are kept."""
    path = make_workbook(tmp_path / "db.xlsx", [["ID", "Volt*", "Lot*"]])
    table, layout = apply_layout(_table(), load_layout(path, "Word layout", required=False))
    assert layout.warnings == []
    assert [(c.name, c.user_input) for c in table.columns] == [
        ("ID", False), ("Name", False), ("Volt", True), ("Lot", True)]
    assert table.items[0].values == {"ID": "P-1", "Name": "Pump", "Volt": "", "Lot": ""}


def test_apply_layout_unknown_column_becomes_blank_with_warning(tmp_path: Path) -> None:
    """A mapped column the database lacks (without '*') is a blank Word column plus a warning."""
    path = make_workbook(tmp_path / "db.xlsx", [["ID", "Colour"]])
    table, layout = apply_layout(_table(), load_layout(path, "Word layout", required=False))
    assert table.column_names == ["ID", "Name", "Volt"]  # not added to the browser table
    assert layout.column_keys == ["ID", None]
    assert layout.column_labels == ["ID", "Colour (blank)"]
    [warning] = layout.warnings
    assert "'Colour' (cell B1) is not in the item database" in warning
    assert "Available columns: ID, Name, Volt" in warning


def test_empty_mapping_cell_and_dash_give_blank_columns(tmp_path: Path) -> None:
    """An empty mapping cell inside the table warns; '-' is blank on purpose, also at the edge."""
    path = make_workbook(tmp_path / "db.xlsx", [["Group", None, None, None], ["-", "Name", None, "Volt"]],
                         ("A1:D1",))
    layout = load_layout(path, "Word layout", required=False)
    assert layout.column_keys == [None, "Name", None, "Volt"]
    assert layout.column_labels == ["(blank)", "Name", "(blank)", "Volt"]
    [warning] = layout.warnings
    assert "Cell C2 in the mapping row is empty" in warning and "'-'" in warning
    assert [(c.text, c.col_span) for c in layout.cells] == [("Group", 4)]


def test_unmapped_database_columns_are_allowed(tmp_path: Path) -> None:
    """Database columns the layout does not name stay in the browser but not in Word, without warnings."""
    path = make_workbook(tmp_path / "db.xlsx", [["Volt"]])
    table, layout = apply_layout(_table(), load_layout(path, "Word layout", required=False))
    assert table.column_names == ["ID", "Name", "Volt"]
    assert layout.column_keys == ["Volt"] and layout.warnings == []


def test_apply_layout_refuses_id_as_input(tmp_path: Path) -> None:
    """The ID column cannot become an input column through the layout."""
    path = make_workbook(tmp_path / "db.xlsx", [["ID*"]])
    with pytest.raises(LayoutError, match="ID column"):
        apply_layout(_table(), load_layout(path, "Word layout", required=False))


def test_data_sheet_skips_layout_sheet(tmp_path: Path) -> None:
    """The layout sheet may come first: the data is read from the first other sheet."""
    path = make_workbook(tmp_path / "db.xlsx", [["Name"]], layout_first=True)
    table = ExcelItemSource(path, skip_sheet="Word layout").load()
    assert [i.id for i in table.items] == ["P-1", "P-2"]


def _span(tc) -> tuple[int, str | None]:
    """Return (gridSpan, vMerge value) of a w:tc element."""
    tc_pr = tc.tcPr
    span = tc_pr.find(qn("w:gridSpan"))
    vmerge = tc_pr.find(qn("w:vMerge"))
    return (int(span.get(qn("w:val"))) if span is not None else 1,
            None if vmerge is None else (vmerge.get(qn("w:val")) or "continue"))


def test_word_table_has_merged_multi_level_header(tmp_path: Path) -> None:
    """The Word table gets the header rows with merges, repeated on each page, then data in mapping order."""
    path = make_workbook(tmp_path / "db.xlsx", SKETCH, SKETCH_MERGES, {"A": 20, "B": 10, "C": 10})
    layout = load_layout(path, "Word layout", required=False)
    items = [Item(id="P-1", values={"Volt": "230", "Amp": "5", "Ref": "R-9", "Name": "Pump"})]
    doc = build_document(TEMPLATE, ["ignored"], items, layout=layout)
    table = doc.tables[-1]
    trs = table._tbl.tr_lst
    assert len(trs) == 3
    row0 = trs[0].tc_lst
    assert [_span(tc) for tc in row0] == [(2, None), (1, "restart")]
    assert [tc.xpath("string(.)") for tc in row0] == ["Electrical", "Doc ref"]
    row1 = trs[1].tc_lst
    assert [_span(tc) for tc in row1] == [(1, None), (1, None), (1, "continue")]
    assert [tc.xpath("string(.)") for tc in trs[2].tc_lst] == ["230", "5", "R-9"]
    for tr in trs[:2]:
        assert tr.trPr.find(qn("w:tblHeader")) is not None
    assert trs[2].trPr.find(qn("w:tblHeader")) is None
    # Widths follow Excel (20:10:10); the merged cell spans the first two.
    grid = [int(c.get(qn("w:w"))) for c in table._tbl.tblGrid.findall(qn("w:gridCol"))]
    assert abs(grid[0] / grid[1] - 2) < 0.01 and abs(grid[1] - grid[2]) <= 1
    merged_w = int(row0[0].tcPr.find(qn("w:tcW")).get(qn("w:w")))
    assert abs(merged_w - (grid[0] + grid[1])) <= 2


def test_header_line_breaks_from_excel_are_kept(tmp_path: Path) -> None:
    """Alt+Enter in an Excel header cell becomes a line break in Word."""
    path = make_workbook(tmp_path / "db.xlsx", [["Rated\nvoltage"], ["Volt"]])
    layout = load_layout(path, "Word layout", required=False)
    doc = build_document(TEMPLATE, [], [Item(id="P-1", values={"Volt": "230"})], layout=layout)
    tc = doc.tables[-1]._tbl.tr_lst[0].tc_lst[0]
    assert tc.findall(".//" + qn("w:br"))


# ---- API ---------------------------------------------------------------------


def _client(tmp_path: Path, db: Path, **cfg) -> TestClient:
    settings = Config(items_file=db, template_file=TEMPLATE, sessions_dir=tmp_path / "s",
                      output_dir=tmp_path / "out", **cfg)
    return TestClient(create_app(settings), client=LOCAL)


def test_api_items_and_export_follow_the_layout(tmp_path: Path) -> None:
    """GET /api/items reports the layout and its input column; export writes the mapped columns."""
    db = make_workbook(tmp_path / "db.xlsx", [["Electrical", None, None], ["Name", "Volt", "Lot*"]], ("A1:C1",))
    c = _client(tmp_path, db)
    body = c.get("/api/items").json()
    assert body["layout"] == "db.xlsx › Word layout"
    assert body["word_columns"] == ["Name", "Volt", "Lot"]
    assert [(col["name"], col["user_input"]) for col in body["columns"]] == [
        ("ID", False), ("Name", False), ("Volt", False), ("Amp", False), ("Lot", True)]
    r = c.post("/api/export", json={"selected_ids": ["P-2"], "file_name": "x",
                                    "user_values": {"P-2": {"Lot": "L7"}}})
    assert r.status_code == 200, r.text
    table = Document(r.json()["path"]).tables[-1]
    rows = [[cell.text for cell in row.cells] for row in table.rows]
    assert rows == [["Electrical"] * 3, ["Fan", "24", "L7"]]


def test_api_without_layout_sheet_uses_automatic_table(tmp_path: Path) -> None:
    """No layout sheet: as before, all shown columns in one header row."""
    c = _client(tmp_path, make_workbook(tmp_path / "db.xlsx"))
    body = c.get("/api/items").json()
    assert body["layout"] == "" and body["word_columns"] == ["ID", "Name", "Volt", "Amp"]


def test_api_layout_error_is_reported(tmp_path: Path) -> None:
    """A broken layout makes the item request fail with a message naming the problem."""
    c = _client(tmp_path, make_workbook(tmp_path / "db.xlsx", [["Name", "Name"]]))
    r = c.get("/api/items")
    assert r.status_code == 500
    assert "Word table layout" in r.json()["detail"] and "mapped twice" in r.json()["detail"]


def test_api_blank_columns_warn_and_export_empty(tmp_path: Path) -> None:
    """Gaps and unknown names are reported as warnings by items and export; their Word cells stay empty."""
    rows = [["Head", None, None, None, None], ["Name", None, "Colour", "-", "Lot*"]]
    c = _client(tmp_path, make_workbook(tmp_path / "db.xlsx", rows, ("A1:E1",)))
    body = c.get("/api/items").json()
    assert body["word_columns"] == ["Name", "(blank)", "Colour (blank)", "(blank)", "Lot"]
    assert len(body["warnings"]) == 2
    assert "Colour" not in [col["name"] for col in body["columns"]]
    r = c.post("/api/export", json={"selected_ids": ["P-1"], "file_name": "x",
                                    "user_values": {"P-1": {"Lot": "L1"}}})
    assert r.status_code == 200, r.text
    assert r.json()["warnings"] == body["warnings"]
    table = Document(r.json()["path"]).tables[-1]
    assert [cell.text for cell in table.rows[-1].cells] == ["Pump", "", "", "", "L1"]


def test_api_separate_layout_file_and_settings(tmp_path: Path) -> None:
    """The layout can live in another workbook/sheet, set through the settings."""
    db = make_workbook(tmp_path / "db.xlsx")
    other = Workbook()
    other.active.title = "Mine"
    layout_rows(other.active, [["Amp", "Name"]])
    other_path = tmp_path / "layout.xlsx"
    other.save(other_path)
    c = _client(tmp_path, db)
    r = c.put("/api/settings", json={"table_layout_file": str(other_path), "table_layout_sheet": "Mine"})
    assert r.status_code == 200, r.text
    assert c.get("/api/items").json()["word_columns"] == ["Amp", "Name"]
    # Wrong sheet name in an explicitly chosen file is an error, not a silent fallback.
    c.put("/api/settings", json={"table_layout_sheet": "Nope"})
    assert "Sheet 'Nope' not found" in c.get("/api/items").json()["detail"]
    # Back to the database workbook (which has no layout sheet): automatic table.
    c.put("/api/settings", json={"table_layout_file": "", "table_layout_sheet": "Word layout"})
    assert c.get("/api/items").json()["layout"] == ""


@pytest.mark.parametrize("body, fragment", [
    ({"table_layout_file": "relative.xlsx"}, "full path"),
    ({"table_layout_file": str(ROOT / "README.md")}, ".xlsx"),
    ({"table_layout_file": str(ROOT / "nope.xlsx")}, "not found"),
    ({"table_layout_sheet": "  "}, "sheet"),
])
def test_api_invalid_layout_settings(tmp_path: Path, body: dict, fragment: str) -> None:
    """Bad layout settings are rejected and change nothing."""
    c = _client(tmp_path, make_workbook(tmp_path / "db.xlsx"))
    before = c.get("/api/settings").json()
    r = c.put("/api/settings", json=body)
    assert r.status_code == 400 and fragment in r.json()["detail"]
    assert c.get("/api/settings").json() == before
