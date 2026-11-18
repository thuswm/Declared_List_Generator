"""Tests for the FastAPI routes."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from fastapi.testclient import TestClient

from dlg.app import DOCX_MEDIA_TYPE, create_app
from dlg.config import Config

ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def client(tmp_path: Path) -> TestClient:
    """A client using the example DB/template and an empty temp sessions folder."""
    settings = Config(
        items_file=ROOT / "data" / "items.xlsx",
        template_file=ROOT / "templates" / "declared_list_template.docx",
        sessions_dir=tmp_path / "sessions",
    )
    return TestClient(create_app(settings))


def test_get_items(client: TestClient) -> None:
    """Items are returned with the database name, ID column and shown columns."""
    body = client.get("/api/items").json()
    assert body["source"] == "items.xlsx"
    assert body["id_column"] == "ID"
    assert [c["name"] for c in body["columns"]] == [
        "ID", "Name", "Description", "Category", "Supplier", "Serial number",
    ]
    assert [c["name"] for c in body["columns"] if c["user_input"]] == ["Serial number"]
    assert len(body["items"]) == 12
    first = body["items"][0]
    assert first["id"] == "DL-001"
    assert first["values"]["Supplier"] == "Acme Fluid"
    assert "#Internal note" not in first["values"]


def test_session_lifecycle(client: TestClient) -> None:
    """Save, list, load, overwrite and delete a session."""
    assert client.get("/api/sessions").json() == []

    saved = client.post("/api/sessions", json={"name": "Proj A", "selected_ids": ["DL-002", "DL-001", "DL-002"]})
    assert saved.status_code == 200
    assert saved.json()["selected_ids"] == ["DL-002", "DL-001"]
    assert saved.json()["source"] == "items.xlsx"

    listing = client.get("/api/sessions").json()
    assert [(s["name"], s["count"], s["configuration_name"]) for s in listing] == [("Proj A", 2, "")]

    assert client.get("/api/sessions/Proj A").json()["selected_ids"] == ["DL-002", "DL-001"]

    client.post("/api/sessions", json={"name": "Proj A", "selected_ids": ["DL-003"]})
    assert client.get("/api/sessions/Proj A").json()["selected_ids"] == ["DL-003"]

    assert client.delete("/api/sessions/Proj A").status_code == 204
    assert client.get("/api/sessions").json() == []


def test_session_errors(client: TestClient) -> None:
    """Unknown sessions give 404; bad or empty names give 400/422."""
    assert client.get("/api/sessions/nope").status_code == 404
    assert client.delete("/api/sessions/nope").status_code == 404
    assert client.post("/api/sessions", json={"name": "..", "selected_ids": []}).status_code == 400
    assert client.post("/api/sessions", json={"name": "", "selected_ids": []}).status_code == 422
    assert client.post("/api/sessions", json={"selected_ids": []}).status_code == 422


def test_session_name_with_slash_roundtrips(client: TestClient) -> None:
    """Names containing "/" or markup characters can be loaded and deleted again."""
    name = "E2E <b>Test</b>"
    assert client.post("/api/sessions", json={"name": name, "selected_ids": ["DL-001"]}).status_code == 200
    assert client.get("/api/sessions/E2E%20%3Cb%3ETest%3C%2Fb%3E").json()["name"] == name
    assert client.delete("/api/sessions/E2E%20%3Cb%3ETest%3C%2Fb%3E").status_code == 204


def test_corrupt_session_file_reports_422(client: TestClient, tmp_path: Path) -> None:
    """A damaged session file yields a clear error rather than a crash."""
    (tmp_path / "sessions").mkdir()
    (tmp_path / "sessions" / "broken.json").write_text("{nope")
    assert client.get("/api/sessions/broken").status_code == 422


def test_export_returns_docx_in_database_order(client: TestClient) -> None:
    """The export is a .docx whose rows follow database order."""
    r = client.post("/api/export", json={"selected_ids": ["DL-005", "DL-002", "DL-009"]})
    assert r.status_code == 200
    assert r.headers["content-type"] == DOCX_MEDIA_TYPE
    assert "declared_list.docx" in r.headers["content-disposition"]
    table = Document(BytesIO(r.content)).tables[0]
    assert [c.text for c in table.rows[0].cells] == [
        "ID", "Name", "Description", "Category", "Supplier", "Serial number",
    ]
    assert [row.cells[0].text for row in table.rows[1:]] == ["DL-002", "DL-005", "DL-009"]


def test_export_errors(client: TestClient) -> None:
    """Empty selections and unknown IDs are rejected with clear messages."""
    assert client.post("/api/export", json={"selected_ids": []}).status_code == 400
    r = client.post("/api/export", json={"selected_ids": ["DL-001", "ZZZ", "AAA"]})
    assert r.status_code == 422
    assert "AAA, ZZZ" in r.json()["detail"]


def test_missing_database_and_template(tmp_path: Path) -> None:
    """Server-side configuration problems surface as 500 with a message."""
    settings = Config(
        items_file=tmp_path / "missing.xlsx",
        template_file=tmp_path / "missing.docx",
        sessions_dir=tmp_path / "s",
    )
    c = TestClient(create_app(settings))
    r = c.get("/api/items")
    assert r.status_code == 500 and "not found" in r.json()["detail"]

    ok = Config(
        items_file=ROOT / "data" / "items.xlsx",
        template_file=tmp_path / "missing.docx",
        sessions_dir=tmp_path / "s",
    )
    r = TestClient(create_app(ok)).post("/api/export", json={"selected_ids": ["DL-001"]})
    assert r.status_code == 500 and "Template not found" in r.json()["detail"]


def test_configured_id_column(tmp_path: Path) -> None:
    """``id_column`` from the configuration selects the key column end to end."""
    from openpyxl import Workbook

    wb = Workbook()
    for row in (["Part No", "Title", "#Hidden"], ["P1", "Widget", "x"], ["P2", "Gadget", "y"]):
        wb.active.append(row)
    wb.save(tmp_path / "parts.xlsx")
    settings = Config(
        items_file=tmp_path / "parts.xlsx",
        template_file=ROOT / "templates" / "declared_list_template.docx",
        sessions_dir=tmp_path / "s",
        id_column="Part No",
    )
    c = TestClient(create_app(settings))
    body = c.get("/api/items").json()
    assert [i["id"] for i in body["items"]] == ["P1", "P2"]
    r = c.post("/api/export", json={"selected_ids": ["P2"]})
    rows = [[cell.text for cell in row.cells] for row in Document(BytesIO(r.content)).tables[0].rows]
    assert rows == [["Part No", "Title"], ["P2", "Gadget"]]

    wrong = Config(items_file=tmp_path / "parts.xlsx", sessions_dir=tmp_path / "s")  # default 'ID'
    r = TestClient(create_app(wrong)).get("/api/items")
    assert r.status_code == 500 and "id_column" in r.json()["detail"]


def test_session_v2_fields_via_api(client: TestClient) -> None:
    """Configuration name and user values are saved, listed and loaded."""
    body = {
        "name": "Rig", "configuration_name": "  Rig 7 build B ", "selected_ids": ["DL-001"],
        "user_values": {"DL-001": {"Serial number": "SN-9"}, "DL-002": {"Serial number": "dropped"}},
    }
    saved = client.post("/api/sessions", json=body).json()
    assert saved["configuration_name"] == "Rig 7 build B"
    assert saved["user_values"] == {"DL-001": {"Serial number": "SN-9"}}
    assert client.get("/api/sessions").json()[0]["configuration_name"] == "Rig 7 build B"
    assert client.get("/api/sessions/Rig").json()["schema_version"] == 2


def test_colliding_session_names_via_api(client: TestClient) -> None:
    """'a/b' and 'a_b' are separate sessions."""
    client.post("/api/sessions", json={"name": "a/b", "selected_ids": ["DL-001"]})
    client.post("/api/sessions", json={"name": "a_b", "selected_ids": ["DL-002"]})
    assert client.get("/api/sessions/a%2Fb").json()["selected_ids"] == ["DL-001"]
    assert client.get("/api/sessions/a_b").json()["selected_ids"] == ["DL-002"]
    assert len(client.get("/api/sessions").json()) == 2


def test_export_includes_typed_values(client: TestClient) -> None:
    """Typed values appear in the input column of the Word table; others stay empty."""
    r = client.post("/api/export", json={
        "selected_ids": ["DL-001", "DL-002"],
        "user_values": {"DL-001": {"Serial number": "SN-0001", "Name": "ignored"}},
    })
    assert r.status_code == 200
    rows = [[c.text for c in row.cells] for row in Document(BytesIO(r.content)).tables[0].rows]
    assert rows[1][1] == "Hydraulic Pump"  # non-input column not overridden
    assert [row[5] for row in rows[1:]] == ["SN-0001", ""]
