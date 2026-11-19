"""Tests for the application version (``dlg.get_version`` and ``GET /api/info``)."""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

import dlg
from dlg.app import create_app
from dlg.config import Config

ROOT = Path(__file__).resolve().parent.parent


def pyproject_version() -> str:
    """Return the version written in the project's ``pyproject.toml``."""
    with (ROOT / "pyproject.toml").open("rb") as fh:
        return tomllib.load(fh)["project"]["version"]


def test_version_comes_from_pyproject() -> None:
    """In a checkout the version is read from pyproject.toml (not stale install metadata)."""
    assert dlg.get_version() == pyproject_version()


def test_api_info_returns_the_version(tmp_path: Path) -> None:
    """The GUI header gets the version from ``GET /api/info``."""
    r = TestClient(create_app(Config(sessions_dir=tmp_path))).get("/api/info")
    assert r.status_code == 200 and r.json() == {"version": pyproject_version()}


def test_falls_back_to_installed_metadata(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """Without a readable pyproject.toml (regular install) the installed metadata is used."""
    monkeypatch.setattr(dlg, "_PYPROJECT", tmp_path / "missing.toml")
    monkeypatch.setattr(dlg, "_installed_version", lambda name: "9.9.9")
    assert dlg.get_version() == "9.9.9"


def test_ignores_a_foreign_pyproject(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    """A pyproject.toml of another project (dlg installed inside it) is not used."""
    other = tmp_path / "pyproject.toml"
    other.write_text('[project]\nname = "other"\nversion = "1.2.3"\n', encoding="utf-8")
    monkeypatch.setattr(dlg, "_PYPROJECT", other)
    monkeypatch.setattr(dlg, "_installed_version", lambda name: "9.9.9")
    assert dlg.get_version() == "9.9.9"
