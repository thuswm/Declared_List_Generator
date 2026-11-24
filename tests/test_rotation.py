"""Tests for rotated header text: read from the Excel cell format, written as Word text direction."""

from __future__ import annotations

from pathlib import Path

import pytest
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Alignment

from dlg.docgen import ROTATED_PADDING_PT, build_document, rotated_text_height
from dlg.layout import load_layout
from dlg.models import Item

ROOT = Path(__file__).resolve().parent.parent
TEMPLATE = ROOT / "templates" / "declared_list_template.docx"
ITEMS = [Item(id="P-1", values={"ID": "P-1", "Name": "Pump", "Volt": "230"})]


def make_layout(path: Path, rows, rotations: dict[str, int], merges=(), heights: dict[int, float] | None = None) -> Path:
    """Write a workbook whose first sheet is the layout, with rotated cells and row heights."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Word layout"
    for row in rows:
        ws.append(row)
    for rng in merges:
        ws.merge_cells(rng)
    for ref, angle in rotations.items():
        ws[ref].alignment = Alignment(textRotation=angle)
    for row, height in (heights or {}).items():
        ws.row_dimensions[row].height = height
    wb.save(path)
    return path


def _header_props(doc, row: int) -> list[tuple[str | None, list[str]]]:
    """Text direction and tcPr child tags of each cell in a header row."""
    out = []
    for tc in doc.tables[0].rows[row]._tr.tc_lst:
        tc_pr = tc.tcPr
        direction = tc_pr.find(qn("w:textDirection"))
        out.append((None if direction is None else direction.get(qn("w:val")),
                    [child.tag.split("}")[1] for child in tc_pr]))
    return out


def _row_height(doc, row: int) -> tuple[float | None, str | None]:
    """Height (points) and rule of a table row."""
    r = doc.tables[0].rows[row]
    return (None if r.height is None else round(r.height.pt, 1),
            None if r.height_rule is None else r.height_rule.name)


def test_rotation_read_from_excel_alignment(tmp_path: Path) -> None:
    """-90° (180) reads down, +90° reads up; merged cells use their top-left cell."""
    path = make_layout(tmp_path / "l.xlsx",
                       [["Group", None, "Serial number"], ["ID", "Name", "Volt"]],
                       {"A1": 90, "C1": 180}, merges=("A1:B1",))
    layout = load_layout(path, "Word layout", required=True)
    assert [(c.text, c.rotation) for c in layout.cells] == [("Group", "up"), ("Serial number", "down")]
    assert layout.warnings == []
    assert layout.row_heights == [None]


def test_rotation_of_mapping_names_when_they_are_the_header(tmp_path: Path) -> None:
    """Without header rows, the mapping cells' own rotation and row height apply."""
    path = make_layout(tmp_path / "l.xlsx", [["ID", "Serial number"]], {"B1": 180}, heights={1: 70})
    layout = load_layout(path, "Word layout", required=True)
    assert [c.rotation for c in layout.cells] == ["", "down"]
    assert layout.row_heights == [70.0]


@pytest.mark.parametrize("angle, fragment", [(45, "45°"), (135, "-45°"), (255, "stacked vertical text")])
def test_other_angles_warn_and_stay_horizontal(tmp_path: Path, angle: int, fragment: str) -> None:
    """Angles Word cannot show give a warning naming the cell."""
    path = make_layout(tmp_path / "l.xlsx", [["Head", "Other"], ["ID", "Name"]], {"A1": angle})
    layout = load_layout(path, "Word layout", required=True)
    assert layout.cells[0].rotation == ""
    [warning] = layout.warnings
    assert "Cell A1" in warning and fragment in warning


def test_word_text_direction_in_schema_order(tmp_path: Path) -> None:
    """tbRl for down, btLr for up, and textDirection sits between shd and vAlign."""
    path = make_layout(tmp_path / "l.xlsx", [["Up", "Plain", "Down"], ["ID", "Name", "Volt"]],
                       {"A1": 90, "C1": 180})
    layout = load_layout(path, "Word layout", required=True)
    doc = build_document(TEMPLATE, [], ITEMS, layout=layout)
    props = _header_props(doc, 0)
    assert [p[0] for p in props] == ["btLr", None, "tbRl"]
    tags = props[0][1]
    assert tags.index("shd") < tags.index("textDirection") < tags.index("vAlign")
    plain = props[1][1]
    assert plain.index("shd") < plain.index("vAlign")  # shading was appended after vAlign before 0.5.0


def test_row_height_estimated_from_longest_line(tmp_path: Path) -> None:
    """A rotated row gets an 'at least' height for its longest line; other rows are untouched."""
    path = make_layout(tmp_path / "l.xlsx",
                       [["Group", None, None], ["ID", "Name", "Serial\nnumber long"], ["ID", "Name", "Volt"]],
                       {"C2": 180}, merges=("A1:C1",))
    layout = load_layout(path, "Word layout", required=True)
    doc = build_document(TEMPLATE, [], ITEMS, layout=layout, font_size=10)
    assert _row_height(doc, 0) == (None, None)
    assert _row_height(doc, 1) == (round(rotated_text_height("Serial\nnumber long", 10), 1), "AT_LEAST")
    assert rotated_text_height("Serial\nnumber long", 10) == pytest.approx(11 * 10 * 0.6 + ROTATED_PADDING_PT)


def test_row_height_from_excel_wins(tmp_path: Path) -> None:
    """A height stored in the layout sheet is used instead of the estimate."""
    path = make_layout(tmp_path / "l.xlsx", [["A very long rotated title", "Name"], ["ID", "Name"]],
                       {"A1": 90}, heights={1: 55})
    layout = load_layout(path, "Word layout", required=True)
    doc = build_document(TEMPLATE, [], ITEMS, layout=layout)
    assert _row_height(doc, 0) == (55.0, "AT_LEAST")


def test_rotated_cell_merged_down_shares_its_height(tmp_path: Path) -> None:
    """A rotated cell spanning two rows splits its estimate between them."""
    path = make_layout(tmp_path / "l.xlsx", [["Long rotated title", "Group"], [None, "Sub"], ["ID", "Name"]],
                       {"A1": 180}, merges=("A1:A2",))
    layout = load_layout(path, "Word layout", required=True)
    doc = build_document(TEMPLATE, [], ITEMS, layout=layout, font_size=8)
    half = round(rotated_text_height("Long rotated title", 8) / 2, 1)
    assert _row_height(doc, 0) == (half, "AT_LEAST")
    assert _row_height(doc, 1) == (half, "AT_LEAST")


def test_layout_endpoint_reports_rotation(tmp_path: Path) -> None:
    """GET /api/layout passes the rotation on to the Check Word table dialog."""
    from fastapi.testclient import TestClient

    from dlg.app import create_app
    from dlg.config import Config

    layout = make_layout(tmp_path / "l.xlsx", [["No.", "Designation"], ["ID", "Name"]], {"B1": 180})
    cfg = Config(items_file=ROOT / "data" / "items.xlsx", template_file=TEMPLATE, sessions_dir=tmp_path / "s",
                 output_dir=tmp_path, table_layout_file=layout)
    body = TestClient(create_app(cfg)).get("/api/layout").json()
    assert [(c["text"], c["rotation"]) for c in body["cells"]] == [("No.", ""), ("Designation", "down")]
