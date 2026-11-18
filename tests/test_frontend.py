"""Tests for the browser GUI: static serving and the JavaScript helper tests."""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

import pytest
from fastapi.testclient import TestClient

from dlg.app import create_app
from dlg.config import Config

ROOT = Path(__file__).resolve().parent.parent


def test_static_files_served_with_module_safe_mime(tmp_path: Path) -> None:
    """The GUI is served at / and scripts use a JavaScript MIME type.

    Browsers refuse ES module scripts served with a non-JavaScript type.
    """
    client = TestClient(create_app(Config(sessions_dir=tmp_path)))
    index = client.get("/")
    assert index.status_code == 200 and "Declared List Generator" in index.text
    for name in ("app.js", "logic.js"):
        r = client.get(f"/{name}")
        assert r.status_code == 200
        assert "javascript" in r.headers["content-type"]
    assert client.get("/style.css").headers["content-type"].startswith("text/css")
    assert client.get("/api/items").status_code == 200  # API not shadowed by static mount


def test_index_declares_an_icon(tmp_path: Path) -> None:
    """The page declares its own icon so browsers don't request (and 404 on) /favicon.ico."""
    client = TestClient(create_app(Config(sessions_dir=tmp_path)))
    assert 'rel="icon"' in client.get("/").text


@pytest.mark.skipif(shutil.which("node") is None, reason="Node.js not installed")
def test_javascript_logic_unit_tests() -> None:
    """Run the Node built-in test runner over tests/js."""
    result = subprocess.run(
        ["node", "--test", "tests/js"], cwd=ROOT, capture_output=True, text=True, timeout=60
    )
    assert result.returncode == 0, result.stdout + result.stderr
