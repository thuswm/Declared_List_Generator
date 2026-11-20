"""Tests for writing a text such as "TBD" into empty input cells (0.4.0 step 4)."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from fastapi.testclient import TestClient
from openpyxl import Workbook

from dlg.app import create_app
from dlg.config import Config, ConfigError, load_config, save_config
from dlg.docgen import fill_empty_inputs
from dlg.models import Item

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "declared_list_template.docx"
LOCAL = ("127.0.0.1", 50000)


def test_fill_empty_inputs_counts_and_keeps_typed_values() -> None:
    """Empty and space-only cells get the text; typed values and other columns are untouched."""
    items = [
        Item(id="1", values={"ID": "1", "Note": "", "Lot": "L1", "Desc": ""}),
        Item(id="2", values={"ID": "2", "Note": "  ", "Lot": ""}),
    ]
    filled, counts = fill_empty_inputs(items, ["Note", "Lot", "Qty"], "TBD")
    assert [i.values for i in filled] == [
        {"ID": "1", "Note": "TBD", "Lot": "L1", "Desc": "", "Qty": "TBD"},
        {"ID": "2", "Note": "TBD", "Lot": "TBD", "Qty": "TBD"},
    ]
    assert counts == {"Note": 2, "Lot": 1, "Qty": 2}
    assert items[0].values["Note"] == ""  # inputs not modified


def test_config_round_trip_and_validation(tmp_path: Path) -> None:
    """The two settings are saved as TOML bool/string and checked on load."""
    path = tmp_path / "config.toml"
    cfg = load_config(path)
    assert (cfg.mark_empty_inputs, cfg.empty_input_text) == (False, "TBD")
    cfg.mark_empty_inputs, cfg.empty_input_text = True, "to be defined"
    save_config(cfg, path)
    text = path.read_text(encoding="utf-8")
    assert "mark_empty_inputs = true" in text
    assert (load_config(path).mark_empty_inputs, load_config(path).empty_input_text) == (True, "to be defined")
    for old, bad, fragment in [("mark_empty_inputs = true", 'mark_empty_inputs = "yes"', "true or false"),
                               ('empty_input_text = "to be defined"', 'empty_input_text = " "', "such as TBD")]:
        path.write_text(text.replace(old, bad), encoding="utf-8")
        with pytest.raises(ConfigError, match=fragment):
            load_config(path)


def _client(tmp_path: Path, layout_rows: list[list] | None) -> TestClient:
    """App over a small database (input column 'Note*', a DB column with an empty cell) and optional layout."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Items"
    for row in [["ID", "Name", "Note*"], ["P-1", "Pump", None], ["P-2", None, None]]:
        ws.append(row)
    if layout_rows is not None:
        layout = wb.create_sheet("Word layout")
        for row in layout_rows:
            layout.append(row)
    db = tmp_path / "db.xlsx"
    wb.save(db)
    cfg = Config(items_file=db, template_file=TEMPLATE, sessions_dir=tmp_path / "s", output_dir=tmp_path / "out")
    return TestClient(create_app(cfg), client=LOCAL)


def _export_rows(c: TestClient, user_values: dict) -> list[list[str]]:
    r = c.post("/api/export", json={"selected_ids": ["P-1", "P-2"], "file_name": "x", "user_values": user_values,
                                    "overwrite": True})
    assert r.status_code == 200, r.text
    return [[cell.text for cell in row.cells] for row in Document(r.json()["path"]).tables[-1].rows]


def test_export_marks_only_input_columns_when_enabled(tmp_path: Path) -> None:
    """Off: empty stays empty. On: empty input cells get the text; empty DB values stay empty."""
    c = _client(tmp_path, None)
    typed = {"P-1": {"Note": "ok"}}
    assert _export_rows(c, typed)[1:] == [["P-1", "Pump", "ok"], ["P-2", "", ""]]
    r = c.put("/api/settings", json={"mark_empty_inputs": True, "empty_input_text": " n/a "})
    assert r.status_code == 200
    assert (r.json()["mark_empty_inputs"], r.json()["empty_input_text"]) == (True, "n/a")
    assert _export_rows(c, typed)[1:] == [["P-1", "Pump", "ok"], ["P-2", "", "n/a"]]


def test_export_with_layout_keeps_blank_columns_empty(tmp_path: Path) -> None:
    """In a layout table, layout-only input columns get the text too; blank columns do not."""
    c = _client(tmp_path, [["ID", "-", "Note*", "Lot*"]])
    c.put("/api/settings", json={"mark_empty_inputs": True})
    assert _export_rows(c, {"P-1": {"Lot": "L1"}})[1:] == [["P-1", "", "TBD", "L1"], ["P-2", "", "TBD", "TBD"]]


def test_api_rejects_empty_text(tmp_path: Path) -> None:
    """An empty text is refused and changes nothing."""
    c = _client(tmp_path, None)
    r = c.put("/api/settings", json={"mark_empty_inputs": True, "empty_input_text": "  "})
    assert r.status_code == 400 and "such as TBD" in r.json()["detail"]
    assert c.get("/api/settings").json()["mark_empty_inputs"] is False
