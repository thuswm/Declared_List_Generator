"""Tests for the command-line launcher."""

from __future__ import annotations

import pytest

from dlg import __main__ as launcher


@pytest.fixture
def calls(monkeypatch: pytest.MonkeyPatch) -> dict:
    """Replace uvicorn.run and threading.Timer with recorders."""
    rec: dict = {"timers": []}
    monkeypatch.setattr(launcher.uvicorn, "run", lambda *a, **k: rec.update(run=(a, k)))

    class FakeTimer:
        """Records the scheduled call instead of starting a thread."""

        def __init__(self, delay: float, fn, args=()) -> None:
            rec["timers"].append((delay, fn, args))

        def start(self) -> None:
            """Pretend to start."""

    monkeypatch.setattr(launcher.threading, "Timer", FakeTimer)
    return rec


def test_defaults_bind_localhost_and_open_browser(calls: dict) -> None:
    """By default the server binds 127.0.0.1:8000 and schedules the browser."""
    launcher.main([])
    assert calls["run"] == (("dlg.app:app",), {"host": "127.0.0.1", "port": 8000})
    assert calls["timers"][0][2] == ("http://localhost:8000/",)


def test_no_browser_and_custom_port(calls: dict) -> None:
    """``--no-browser`` skips the browser; ``--port`` is honoured."""
    launcher.main(["--no-browser", "--port", "9001"])
    assert calls["timers"] == []
    assert calls["run"][1]["port"] == 9001
