"""Tests for opening an exported document or its folder (``POST /api/open``, ``dlg.opener``)."""

from __future__ import annotations

import subprocess
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from dlg import opener
from dlg.app import create_app
from dlg.config import Config
from dlg.opener import OpenError, open_path

ROOT = Path(__file__).resolve().parent.parent
LOCAL = ("127.0.0.1", 50000)


@pytest.fixture
def setup(tmp_path: Path) -> tuple[TestClient, list, Path]:
    """A local client with a recording opener and an output folder holding one document."""
    calls: list[tuple[Path, bool]] = []
    out = tmp_path / "out"
    out.mkdir()
    (out / "list.docx").write_bytes(b"x")
    (out / "notes.txt").write_text("x")
    app = create_app(Config(sessions_dir=tmp_path / "s", output_dir=out),
                     opener=lambda path, folder: calls.append((path, folder)))
    return TestClient(app, client=LOCAL), calls, out


def test_opens_document_and_folder(setup: tuple[TestClient, list, Path]) -> None:
    """The document in the output folder is handed to the opener, as file or folder."""
    c, calls, out = setup
    assert c.post("/api/open", json={"file_name": "list.docx"}).status_code == 204
    assert c.post("/api/open", json={"file_name": "list.docx", "folder": True}).status_code == 204
    assert calls == [(out / "list.docx", False), (out / "list.docx", True)]


@pytest.mark.parametrize("name, status", [
    ("../list.docx", 400),
    ("sub/list.docx", 400),
    ("..\\\\list.docx", 400),
    ("notes.txt", 400),
    (str(ROOT / "README.md"), 400),
    ("missing.docx", 404),
])
def test_refuses_anything_but_an_existing_document_in_the_output_folder(
    setup: tuple[TestClient, list, Path], name: str, status: int
) -> None:
    """Paths, other file types and missing files are refused; nothing is opened."""
    c, calls, _ = setup
    assert c.post("/api/open", json={"file_name": name}).status_code == status
    assert calls == []


def test_is_local_only(setup: tuple[TestClient, list, Path]) -> None:
    """Another computer cannot make this one open programs."""
    c, calls, _ = setup
    remote = TestClient(c.app, client=("192.168.1.20", 50000))
    assert remote.post("/api/open", json={"file_name": "list.docx"}).status_code == 403
    assert calls == []


def test_opener_error_is_reported(tmp_path: Path) -> None:
    """A failing opener gives a readable 500."""
    out = tmp_path / "out"
    out.mkdir()
    (out / "a.docx").write_bytes(b"x")

    def broken(path: Path, folder: bool) -> None:
        raise OpenError("Cannot open a.docx: no program")

    c = TestClient(create_app(Config(sessions_dir=tmp_path, output_dir=out), opener=broken), client=LOCAL)
    r = c.post("/api/open", json={"file_name": "a.docx"})
    assert r.status_code == 500 and "no program" in r.json()["detail"]


@pytest.mark.parametrize("platform, folder, expected", [
    ("linux", False, ["xdg-open", "{file}"]),
    ("linux", True, ["xdg-open", "{dir}"]),
    ("darwin", False, ["open", "{file}"]),
    ("darwin", True, ["open", "-R", "{file}"]),
    ("win32", True, ["explorer", "/select,{file}"]),
])
def test_platform_commands(tmp_path: Path, monkeypatch: pytest.MonkeyPatch, platform: str, folder: bool,
                           expected: list[str]) -> None:
    """Each platform gets its usual command (nothing is actually started)."""
    started: list[list[str]] = []
    monkeypatch.setattr(subprocess, "Popen", lambda args, **kw: started.append(args))
    path = tmp_path / "a.docx"
    open_path(path, folder, platform=platform)
    assert started == [[a.format(file=path, dir=tmp_path) for a in expected]]


def test_windows_document_uses_startfile(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """On Windows the document opens with its registered program (os.startfile)."""
    opened: list[Path] = []
    monkeypatch.setattr(opener.os, "startfile", opened.append, raising=False)
    open_path(tmp_path / "a.docx", False, platform="win32")
    assert opened == [tmp_path / "a.docx"]


def test_missing_program_becomes_open_error(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """If e.g. xdg-open is not installed, the user gets a message instead of a crash."""
    def missing(*args, **kwargs):
        raise FileNotFoundError(2, "No such file or directory")

    monkeypatch.setattr(subprocess, "Popen", missing)
    with pytest.raises(OpenError, match="No such file"):
        open_path(tmp_path / "a.docx", False, platform="linux")


@pytest.fixture
def settings_setup(tmp_path: Path) -> tuple[TestClient, list, Config]:
    """A local client with a recording opener and settings pointing at real files."""
    calls: list[tuple[Path, bool]] = []
    cfg = Config(
        items_file=ROOT / "data" / "items.xlsx",
        template_file=ROOT / "templates" / "declared_list_template.docx",
        sessions_dir=tmp_path / "s",
        output_dir=tmp_path,
    )
    app = create_app(cfg, opener=lambda path, folder: calls.append((path, folder)))
    return TestClient(app, client=LOCAL), calls, cfg


def test_open_setting_files_and_folder(settings_setup: tuple[TestClient, list, Config]) -> None:
    """Each saved setting opens its own path; no layout file means the item database."""
    c, calls, cfg = settings_setup
    for setting in ("items_file", "template_file", "table_layout_file", "output_dir"):
        assert c.post("/api/open-setting", json={"setting": setting}).status_code == 204
    assert calls == [(cfg.items_file, False), (cfg.template_file, False),
                     (cfg.items_file, False), (cfg.output_dir, False)]
    cfg.table_layout_file = ROOT / "data" / "layout_example.xlsx"
    assert c.post("/api/open-setting", json={"setting": "table_layout_file"}).status_code == 204
    assert calls[-1] == (cfg.table_layout_file, False)


def test_open_setting_refuses_unknown_missing_and_remote(
    settings_setup: tuple[TestClient, list, Config], tmp_path: Path
) -> None:
    """Only the four settings, only existing paths, only from this computer."""
    c, calls, cfg = settings_setup
    assert c.post("/api/open-setting", json={"setting": "sessions_dir"}).status_code == 422
    cfg.template_file = tmp_path / "gone.docx"
    r = c.post("/api/open-setting", json={"setting": "template_file"})
    assert r.status_code == 404 and "gone.docx" in r.json()["detail"]
    remote = TestClient(c.app, client=("192.168.1.20", 50000))
    assert remote.post("/api/open-setting", json={"setting": "items_file"}).status_code == 403
    assert calls == []


def test_open_setting_reports_opener_failure(tmp_path: Path) -> None:
    """An opener error becomes a readable 500."""
    def broken(path: Path, folder: bool) -> None:
        raise OpenError("Cannot open items.xlsx: no program")

    app = create_app(Config(items_file=ROOT / "data" / "items.xlsx", sessions_dir=tmp_path / "s"), opener=broken)
    r = TestClient(app, client=LOCAL).post("/api/open-setting", json={"setting": "items_file"})
    assert r.status_code == 500 and "no program" in r.json()["detail"]
