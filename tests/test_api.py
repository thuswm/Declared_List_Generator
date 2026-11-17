"""Tests for the FastAPI routes."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pytest
from docx import Document
from fastapi.testclient import TestClient

from dlg.app import DOCX_MEDIA_TYPE, Settings, create_app

ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def client(tmp_path: Path) -> TestClient:
    """A client using the example DB/template and an empty temp sessions folder."""
    settings = Settings(
        items_path=ROOT / "data" / "items.xlsx",
        template_path=ROOT / "templates" / "declared_list_template.docx",
        sessions_dir=tmp_path / "sessions",
    )
    return TestClient(create_app(settings))


def test_get_items(client: TestClient) -> None:
    """Items are returned with the database name."""
    body = client.get("/api/items").json()
    assert body["source"] == "items.xlsx"
    assert len(body["items"]) == 12
    assert body["items"][0]["id"] == "DL-001"


def test_session_lifecycle(client: TestClient) -> None:
    """Save, list, load, overwrite and delete a session."""
    assert client.get("/api/sessions").json() == []

    saved = client.post("/api/sessions", json={"name": "Proj A", "selected_ids": ["DL-002", "DL-001", "DL-002"]})
    assert saved.status_code == 200
    assert saved.json()["selected_ids"] == ["DL-002", "DL-001"]
    assert saved.json()["source"] == "items.xlsx"

    listing = client.get("/api/sessions").json()
    assert [(s["name"], s["count"]) for s in listing] == [("Proj A", 2)]

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
    assert [row.cells[1].text for row in table.rows[1:]] == ["DL-002", "DL-005", "DL-009"]


def test_export_errors(client: TestClient) -> None:
    """Empty selections and unknown IDs are rejected with clear messages."""
    assert client.post("/api/export", json={"selected_ids": []}).status_code == 400
    r = client.post("/api/export", json={"selected_ids": ["DL-001", "ZZZ", "AAA"]})
    assert r.status_code == 422
    assert "AAA, ZZZ" in r.json()["detail"]


def test_missing_database_and_template(tmp_path: Path) -> None:
    """Server-side configuration problems surface as 500 with a message."""
    settings = Settings(
        items_path=tmp_path / "missing.xlsx",
        template_path=tmp_path / "missing.docx",
        sessions_dir=tmp_path / "s",
    )
    c = TestClient(create_app(settings))
    r = c.get("/api/items")
    assert r.status_code == 500 and "not found" in r.json()["detail"]

    ok = Settings(
        items_path=ROOT / "data" / "items.xlsx",
        template_path=tmp_path / "missing.docx",
        sessions_dir=tmp_path / "s",
    )
    r = TestClient(create_app(ok)).post("/api/export", json={"selected_ids": ["DL-001"]})
    assert r.status_code == 500 and "Template not found" in r.json()["detail"]


def test_settings_read_environment(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    """Environment variables override the default file locations."""
    monkeypatch.setenv("DLG_ITEMS", str(tmp_path / "i.xlsx"))
    monkeypatch.setenv("DLG_TEMPLATE", str(tmp_path / "t.docx"))
    monkeypatch.setenv("DLG_SESSIONS", str(tmp_path / "s"))
    s = Settings()
    assert (s.items_path, s.template_path, s.sessions_dir) == (
        tmp_path / "i.xlsx", tmp_path / "t.docx", tmp_path / "s",
    )
