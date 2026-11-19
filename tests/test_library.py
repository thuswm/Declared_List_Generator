"""Tests for the session library API (``/api/library``) and the session folder settings."""

from __future__ import annotations

import json
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from dlg.app import create_app
from dlg.config import Config, load_config, save_config
from dlg.models import Session
from dlg.sessions import SessionStore

ROOT = Path(__file__).resolve().parent.parent
LOCAL = ("127.0.0.1", 50000)


@pytest.fixture
def env(tmp_path: Path) -> tuple[TestClient, Path, Path]:
    """A local client with a user library, a shared library folder and a saved config file."""
    user, shared = tmp_path / "user", tmp_path / "shared"
    user.mkdir()
    shared.mkdir()
    cfg_path = tmp_path / "config.toml"
    settings = Config(items_file=ROOT / "data" / "items.xlsx", sessions_dir=user, output_dir=tmp_path)
    save_config(settings, cfg_path)
    return TestClient(create_app(settings, cfg_path), client=LOCAL), user, shared


def test_lists_user_library_with_files_and_skipped(env: tuple[TestClient, Path, Path]) -> None:
    """Default folder is sessions_dir; entries carry file and path; broken files are reported."""
    c, user, _ = env
    SessionStore(user).save(Session(name="A/B", configuration_name="Rig 7", selected_ids=["P-1", "P-2"]))
    (user / "broken.json").write_text("{not json", encoding="utf-8")
    body = c.get("/api/library").json()
    assert body["folder"] == str(user)
    [entry] = body["sessions"]
    assert entry["name"] == "A/B" and entry["configuration_name"] == "Rig 7" and entry["count"] == 2
    assert entry["file"] == "A_B.json" and entry["path"] == str(user / "A_B.json")
    assert [s["file"] for s in body["skipped"]] == ["broken.json"]
    assert "broken.json" in body["skipped"][0]["error"]


def test_lists_any_folder_newest_first(env: tuple[TestClient, Path, Path]) -> None:
    """Another folder can be listed; sessions come newest first."""
    c, _, shared = env
    store = SessionStore(shared)
    store.save(Session(name="old"))
    store.save(Session(name="new"))
    body = c.get("/api/library", params={"folder": str(shared)}).json()
    assert [s["name"] for s in body["sessions"]] == ["new", "old"]


def test_missing_user_library_is_empty_but_other_missing_folder_is_404(tmp_path: Path) -> None:
    """sessions_dir is created on first save, so its absence is not an error; other folders must exist."""
    c = TestClient(create_app(Config(sessions_dir=tmp_path / "not yet")), client=LOCAL)
    assert c.get("/api/library").json()["sessions"] == []
    assert c.get("/api/library", params={"folder": str(tmp_path / "nope")}).status_code == 404
    assert c.get("/api/library", params={"folder": "relative"}).status_code == 400


def test_load_and_delete_by_path(env: tuple[TestClient, Path, Path]) -> None:
    """A session file in any folder is loaded and deleted by its path."""
    c, _, shared = env
    path = SessionStore(shared).save(Session(name="Shared one", selected_ids=["P-1"]))
    r = c.get("/api/library/session", params={"path": str(path)})
    assert r.status_code == 200 and r.json()["name"] == "Shared one"
    assert c.delete("/api/library/session", params={"path": str(path)}).status_code == 204
    assert not path.exists()
    assert c.get("/api/library/session", params={"path": str(path)}).status_code == 404


def test_delete_refuses_files_that_are_not_sessions(env: tuple[TestClient, Path, Path], tmp_path: Path) -> None:
    """A wrong path can never delete an unrelated file."""
    c, _, shared = env
    other = tmp_path / "notes.txt"
    other.write_text("keep me", encoding="utf-8")
    fake = shared / "data.json"
    fake.write_text(json.dumps([1, 2, 3]), encoding="utf-8")
    assert c.delete("/api/library/session", params={"path": str(other)}).status_code == 400
    assert c.delete("/api/library/session", params={"path": str(fake)}).status_code == 400
    assert other.exists() and fake.exists()
    assert c.get("/api/library/session", params={"path": str(fake)}).status_code == 422


def test_library_is_local_only(env: tuple[TestClient, Path, Path]) -> None:
    """Listing, reading and deleting files by path is refused for other computers."""
    c, user, _ = env
    remote = TestClient(c.app, client=("192.168.1.20", 50000))
    assert remote.get("/api/library").status_code == 403
    assert remote.get("/api/library/session", params={"path": str(user / "x.json")}).status_code == 403
    assert remote.delete("/api/library/session", params={"path": str(user / "x.json")}).status_code == 403


def test_change_session_folders(env: tuple[TestClient, Path, Path], tmp_path: Path) -> None:
    """Both library folders can be set (and the shared one unset); they persist and apply immediately."""
    c, _, shared = env
    other_user = tmp_path / "other user"
    other_user.mkdir()
    r = c.put("/api/settings", json={"sessions_dir": str(other_user), "global_sessions_dir": str(shared)})
    assert r.status_code == 200, r.text
    assert r.json()["sessions_dir"] == str(other_user) and r.json()["global_sessions_dir"] == str(shared)
    saved = load_config(tmp_path / "config.toml")
    assert saved.sessions_dir == other_user and saved.global_sessions_dir == shared
    # New user library is used for saving and listing straight away.
    assert c.post("/api/sessions", json={"name": "n", "selected_ids": []}).status_code == 200
    assert (other_user / "n.json").exists()
    assert [s["name"] for s in c.get("/api/sessions").json()] == ["n"]
    # Unset the shared library.
    assert c.put("/api/settings", json={"global_sessions_dir": ""}).json()["global_sessions_dir"] == ""
    assert load_config(tmp_path / "config.toml").global_sessions_dir is None


@pytest.mark.parametrize("body, fragment", [
    ({"sessions_dir": "relative"}, "full path"),
    ({"sessions_dir": str(ROOT / "no" / "such" / "folder")}, "My sessions folder not found"),
    ({"global_sessions_dir": str(ROOT / "no" / "such" / "folder")}, "Shared library folder not found"),
])
def test_invalid_session_folders_change_nothing(env: tuple[TestClient, Path, Path], body: dict, fragment: str) -> None:
    """A bad folder is rejected with a message and nothing changes."""
    c, _, _ = env
    before = c.get("/api/settings").json()
    r = c.put("/api/settings", json=body)
    assert r.status_code == 400 and fragment in r.json()["detail"]
    assert c.get("/api/settings").json() == before
