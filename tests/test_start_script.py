"""Checks of the Windows start script ``start_dlg.bat`` (it cannot be run on Linux)."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SCRIPT = ROOT / "start_dlg.bat"


def test_start_script_uses_crlf_and_plain_ascii() -> None:
    """cmd.exe reads batch files in the console code page and needs CRLF for reliable labels."""
    data = SCRIPT.read_bytes()
    data.decode("ascii")  # raises on any non-ASCII character
    assert b"\n" not in data.replace(b"\r\n", b"")


def test_start_script_runs_the_venv_python_from_its_own_folder() -> None:
    """It changes to its folder, uses the project's .venv Python, passes arguments on and pauses on errors."""
    text = SCRIPT.read_bytes().decode("ascii")  # not read_text: it would turn CRLF into \n
    assert 'cd /d "%~dp0"' in text
    assert r'set "DLG_PYTHON=%~dp0.venv\Scripts\python.exe"' in text
    assert '"%DLG_PYTHON%" -m dlg %*' in text
    assert text.count("pause") == 2
    for label in (":no_venv", ":failed"):
        assert f"goto {label}" in text and f"\r\n{label}\r\n" in text


def test_gitattributes_keeps_crlf_for_batch_files() -> None:
    """Without this, a checkout with LF endings could break the script on Windows."""
    assert "*.bat text eol=crlf" in (ROOT / ".gitattributes").read_text(encoding="utf-8")
