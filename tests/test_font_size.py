"""Tests for the Word table font size setting (0.3.0 step 6)."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document
from docx.oxml.ns import qn
from fastapi.testclient import TestClient

from dlg.app import create_app
from dlg.config import Config, ConfigError, load_config, save_config
from dlg.docgen import build_document
from dlg.layout import HeaderCell, LayoutColumn, TableLayout
from dlg.models import Item

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "declared_list_template.docx"
LOCAL = ("127.0.0.1", 50000)
ITEMS = [Item(id="P-1", values={"ID": "P-1", "Name": "Pump", "Note": ""})]


def _sizes(table) -> tuple[set[str], set[str]]:
    """Return the w:sz values of all runs and of all paragraph marks in a table."""
    runs = {sz.get(qn("w:val")) for sz in table._tbl.xpath(".//w:r/w:rPr/w:sz")}
    marks = {sz.get(qn("w:val")) for sz in table._tbl.xpath(".//w:pPr/w:rPr/w:sz")}
    return runs, marks


def test_font_size_applies_to_every_cell_including_empty_ones() -> None:
    """Header and data runs and every paragraph mark get the size (empty cells too)."""
    doc = build_document(TEMPLATE, ["ID", "Name", "Note"], ITEMS, font_size=8.5)
    table = doc.tables[-1]
    runs, marks = _sizes(table)
    assert runs == {"17"} and marks == {"17"}  # half-points
    n_paragraphs = len(table._tbl.xpath(".//w:p"))
    assert len(table._tbl.xpath(".//w:pPr/w:rPr/w:sz")) == n_paragraphs


def test_font_size_with_merged_header_and_valid_order() -> None:
    """Merged header cells get the size once; rPr follows jc inside pPr."""
    layout = TableLayout(
        columns=[LayoutColumn(name="ID"), LayoutColumn(name="Name")], header_rows=1,
        cells=[HeaderCell(0, 0, "Item", 1, 2)], widths=[1, 1],
    )
    doc = build_document(TEMPLATE, [], ITEMS, layout=layout, font_size=7)
    table = doc.tables[-1]
    assert _sizes(table) == ({"14"}, {"14"})
    p_pr = table._tbl.tr_lst[0].tc_lst[0].xpath(".//w:pPr")[0]
    assert [child.tag.split("}")[1] for child in p_pr] == ["spacing", "jc", "rPr"]


def test_zero_keeps_template_font() -> None:
    """0 = no size and no spacing set anywhere in the table."""
    doc = build_document(TEMPLATE, ["ID", "Name"], ITEMS, font_size=0)
    assert _sizes(doc.tables[-1]) == (set(), set())
    assert not doc.tables[-1]._tbl.xpath(".//w:pPr/w:spacing")


def test_font_size_makes_cell_spacing_compact() -> None:
    """With a size set, every cell paragraph has no space before/after and single line spacing."""
    doc = build_document(TEMPLATE, ["ID", "Name", "Note"], ITEMS, font_size=8)
    tbl = doc.tables[-1]._tbl
    spacings = tbl.xpath(".//w:pPr/w:spacing")
    assert len(spacings) == len(tbl.xpath(".//w:p"))
    for sp in spacings:
        assert (sp.get(qn("w:before")), sp.get(qn("w:after")), sp.get(qn("w:line"))) == ("0", "0", "240")


def test_config_round_trip_and_validation(tmp_path: Path) -> None:
    """The size is stored as a number; bad values in the file are reported."""
    path = tmp_path / "config.toml"
    cfg = load_config(path)
    assert cfg.table_font_size == 0
    cfg.table_font_size = 8.5
    save_config(cfg, path)
    assert load_config(path).table_font_size == 8.5
    text = path.read_text(encoding="utf-8")
    path.write_text(text.replace("table_font_size = 8.5", "table_font_size = 9"), encoding="utf-8")
    assert load_config(path).table_font_size == 9.0  # TOML integer accepted
    for bad in ('"8"', "3", "8.3", "true"):
        path.write_text(text.replace("table_font_size = 8.5", f"table_font_size = {bad}"), encoding="utf-8")
        with pytest.raises(ConfigError, match="table_font_size|Table font size"):
            load_config(path)


def test_api_setting_is_used_by_export(tmp_path: Path) -> None:
    """The size set in Settings is written into the exported table."""
    c = TestClient(create_app(Config(items_file=ROOT / "data" / "items.xlsx", template_file=TEMPLATE,
                                     sessions_dir=tmp_path / "s", output_dir=tmp_path)), client=LOCAL)
    r = c.put("/api/settings", json={"table_font_size": 8})
    assert r.status_code == 200 and r.json()["table_font_size"] == 8
    out = c.post("/api/export", json={"selected_ids": ["DL-001"], "file_name": "f"}).json()["path"]
    assert _sizes(Document(out).tables[-1]) == ({"16"}, {"16"})


@pytest.mark.parametrize("size", [3, 73, 8.25, -1])
def test_api_rejects_bad_sizes(tmp_path: Path, size: float) -> None:
    """Sizes outside 4–72 or not in half points are refused and change nothing."""
    c = TestClient(create_app(Config(sessions_dir=tmp_path / "s")), client=LOCAL)
    r = c.put("/api/settings", json={"table_font_size": size})
    assert r.status_code == 400 and "steps of 0.5" in r.json()["detail"]
    assert c.get("/api/settings").json()["table_font_size"] == 0
