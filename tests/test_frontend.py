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


def test_responses_tell_the_browser_to_revalidate(tmp_path: Path) -> None:
    """GUI files and API answers carry ``Cache-Control: no-cache``.

    Without it browsers kept running an old ``app.js`` after an update until F5.
    """
    client = TestClient(create_app(Config(sessions_dir=tmp_path)))
    for path in ("/", "/app.js", "/logic.js", "/style.css", "/api/info"):
        assert client.get(path).headers["cache-control"] == "no-cache", path


def test_static_files_still_answer_conditional_requests(tmp_path: Path) -> None:
    """Revalidation stays cheap: an unchanged file answers 304."""
    client = TestClient(create_app(Config(sessions_dir=tmp_path)))
    etag = client.get("/app.js").headers["etag"]
    assert client.get("/app.js", headers={"If-None-Match": etag}).status_code == 304


def test_header_has_version_and_file_name_has_reset(tmp_path: Path) -> None:
    """The page has the elements app.js fills/binds for step 1 of 0.3.0."""
    html = TestClient(create_app(Config(sessions_dir=tmp_path))).get("/").text
    assert 'id="version"' in html and 'id="file-name-reset"' in html
