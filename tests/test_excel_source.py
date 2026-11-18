"""Tests for the Excel item source and data models."""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path

import pytest
from openpyxl import Workbook

from dlg.models import Session
from dlg.sources.excel import ExcelItemSource, ItemSourceError, cell_text


def _write(path: Path, rows: list[list]) -> Path:
    """Create a workbook with the given rows on the first sheet."""
    wb = Workbook()
    for row in rows:
        wb.active.append(row)
    wb.save(path)
    return path


def test_columns_come_from_the_sheet_in_order(tmp_path: Path) -> None:
    """All headed columns are shown, in sheet order, with every value as text."""
    p = _write(tmp_path / "a.xlsx", [
        ["Part", "ID", "Weight", "Supplier"],
        ["Pump", "A1", 12, None],
    ])
    table = ExcelItemSource(p).load()
    assert table.column_names == ["Part", "ID", "Weight", "Supplier"]
    assert table.id_column == "ID"
    assert table.items[0].id == "A1"
    assert table.items[0].values == {"Part": "Pump", "ID": "A1", "Weight": "12", "Supplier": ""}
    assert list(table.items[0].values) == table.column_names


def test_omit_marker_and_empty_headers_are_ignored(tmp_path: Path) -> None:
    """Columns headed '#...' or with no header are not shown at all."""
    p = _write(tmp_path / "a.xlsx", [
        ["ID", "#Internal", None, "Name", "  #Also hidden "],
        ["A1", "secret", "orphan", "Pump", "x"],
    ])
    table = ExcelItemSource(p).load()
    assert table.column_names == ["ID", "Name"]
    assert table.items[0].values == {"ID": "A1", "Name": "Pump"}


def test_configurable_id_column(tmp_path: Path) -> None:
    """The ID column is the one named by ``id_column``, wherever it is."""
    p = _write(tmp_path / "a.xlsx", [["Title", "Part No"], ["Widget", "P1"], ["Gadget", "P2"]])
    table = ExcelItemSource(p, id_column="Part No").load()
    assert [i.id for i in table.items] == ["P1", "P2"]
    assert table.column_names == ["Title", "Part No"]


def test_missing_id_column_names_setting_and_columns(tmp_path: Path) -> None:
    """A missing ID column gives an error that says how to fix it."""
    p = _write(tmp_path / "a.xlsx", [["Foo", "Name", "#ID"], ["x", "y", "z"]])
    with pytest.raises(ItemSourceError, match="ID column 'ID' not found") as exc:
        ExcelItemSource(p).load()
    assert "id_column" in str(exc.value) and "Foo, Name" in str(exc.value)


def test_numeric_ids_become_strings_and_blank_rows_skipped(tmp_path: Path) -> None:
    """Numeric IDs are stringified; rows without an ID are ignored; short rows are padded."""
    p = _write(tmp_path / "a.xlsx", [
        ["ID", "Name", "Note"], [1, "One"], [None, "Spacer"], [2.0, "Two", "n"],
    ])
    items = ExcelItemSource(p).load().items
    assert [i.id for i in items] == ["1", "2"]
    assert items[0].values["Note"] == ""


def test_duplicate_ids(tmp_path: Path) -> None:
    """Duplicate IDs are rejected."""
    p = _write(tmp_path / "a.xlsx", [["ID", "Name"], ["a", "x"], ["a", "y"]])
    with pytest.raises(ItemSourceError, match="Duplicate"):
        ExcelItemSource(p).load()


def test_duplicate_headers(tmp_path: Path) -> None:
    """Two shown columns with the same header are rejected (values would collide)."""
    p = _write(tmp_path / "a.xlsx", [["ID", "Name", "Name "], ["a", "x", "y"]])
    with pytest.raises(ItemSourceError, match="more than once"):
        ExcelItemSource(p).load()


def test_empty_sheet(tmp_path: Path) -> None:
    """A workbook without a header row is reported clearly."""
    p = _write(tmp_path / "a.xlsx", [])
    with pytest.raises(ItemSourceError, match="no header"):
        ExcelItemSource(p).load()


def test_missing_file(tmp_path: Path) -> None:
    """A nonexistent file raises ItemSourceError."""
    with pytest.raises(ItemSourceError, match="not found"):
        ExcelItemSource(tmp_path / "nope.xlsx").load()


def test_not_an_xlsx(tmp_path: Path) -> None:
    """A corrupt file raises ItemSourceError rather than a raw library error."""
    bad = tmp_path / "bad.xlsx"
    bad.write_text("not a workbook")
    with pytest.raises(ItemSourceError, match="Cannot read"):
        ExcelItemSource(bad).load()


def test_cell_text_formats() -> None:
    """Dates, whole floats and padded strings are shown sensibly."""
    assert cell_text(None) == ""
    assert cell_text(datetime(2026, 11, 18)) == "2026-11-18"
    assert cell_text(datetime(2026, 11, 18, 9, 30)) == "2026-11-18 09:30:00"
    assert cell_text(date(2026, 1, 2)) == "2026-01-02"
    assert cell_text(3.0) == "3"
    assert cell_text(2.5) == "2.5"
    assert cell_text("  x ") == "x"


def test_session_defaults() -> None:
    """A new Session has the current schema version and timestamps."""
    s = Session(name="x")
    assert s.schema_version == 1
    assert s.selected_ids == []
    assert s.created.tzinfo is not None


def test_example_database_loads() -> None:
    """The example database loads; its '#Internal note' column is omitted."""
    path = Path(__file__).resolve().parent.parent / "data" / "items.xlsx"
    table = ExcelItemSource(path).load()
    assert table.column_names == ["ID", "Name", "Description", "Category", "Supplier"]
    assert len(table.items) == 12
    assert len({i.id for i in table.items}) == 12
    assert table.items[0].values["Supplier"] == "Acme Fluid"
