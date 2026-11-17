"""Tests for the Excel item source and data models."""

from __future__ import annotations

from pathlib import Path

import pytest
from openpyxl import Workbook

from dlg.models import Session
from dlg.sources.excel import ExcelItemSource, ItemSourceError


def _write(path: Path, rows: list[list]) -> Path:
    """Create a workbook with the given rows on the first sheet."""
    wb = Workbook()
    for row in rows:
        wb.active.append(row)
    wb.save(path)
    return path


def test_loads_items_and_extra_columns(tmp_path: Path) -> None:
    """Mapped fields populate the Item; other columns go to ``extra``."""
    p = _write(tmp_path / "a.xlsx", [
        ["ID", "Name", "Description", "Category", "Supplier"],
        ["A1", "Pump", "Desc", "Hyd", "Acme"],
    ])
    items = ExcelItemSource(p).load_items()
    assert len(items) == 1
    assert items[0].id == "A1"
    assert items[0].category == "Hyd"
    assert items[0].extra == {"Supplier": "Acme"}


def test_numeric_ids_become_strings_and_blank_rows_skipped(tmp_path: Path) -> None:
    """Numeric IDs are stringified; rows without an ID are ignored."""
    p = _write(tmp_path / "a.xlsx", [
        ["ID", "Name"], [1, "One"], [None, "Spacer"], [2, "Two"],
    ])
    assert [i.id for i in ExcelItemSource(p).load_items()] == ["1", "2"]


def test_missing_required_column(tmp_path: Path) -> None:
    """A missing ID/Name header raises ItemSourceError."""
    p = _write(tmp_path / "a.xlsx", [["Foo", "Name"], ["x", "y"]])
    with pytest.raises(ItemSourceError, match="ID"):
        ExcelItemSource(p).load_items()


def test_duplicate_ids(tmp_path: Path) -> None:
    """Duplicate IDs are rejected."""
    p = _write(tmp_path / "a.xlsx", [["ID", "Name"], ["a", "x"], ["a", "y"]])
    with pytest.raises(ItemSourceError, match="Duplicate"):
        ExcelItemSource(p).load_items()


def test_missing_file(tmp_path: Path) -> None:
    """A nonexistent file raises ItemSourceError."""
    with pytest.raises(ItemSourceError, match="not found"):
        ExcelItemSource(tmp_path / "nope.xlsx").load_items()


def test_not_an_xlsx(tmp_path: Path) -> None:
    """A corrupt file raises ItemSourceError rather than a raw library error."""
    bad = tmp_path / "bad.xlsx"
    bad.write_text("not a workbook")
    with pytest.raises(ItemSourceError, match="Cannot read"):
        ExcelItemSource(bad).load_items()


def test_custom_column_map(tmp_path: Path) -> None:
    """A custom header mapping is honoured."""
    p = _write(tmp_path / "a.xlsx", [["Part No", "Title"], ["P1", "Widget"]])
    src = ExcelItemSource(p, columns={"id": "Part No", "name": "Title"})
    assert src.load_items()[0].name == "Widget"


def test_session_defaults() -> None:
    """A new Session has the current schema version and timestamps."""
    s = Session(name="x")
    assert s.schema_version == 1
    assert s.selected_ids == []
    assert s.created.tzinfo is not None


def test_example_database_loads() -> None:
    """The generated example database loads with unique, populated items."""
    path = Path(__file__).resolve().parent.parent / "data" / "items.xlsx"
    items = ExcelItemSource(path).load_items()
    assert len(items) == 12
    assert len({i.id for i in items}) == 12
    assert all(i.name and i.category for i in items)
    assert items[0].extra == {"Supplier": "Acme Fluid"}
