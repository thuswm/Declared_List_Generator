"""Tests for the command-line launcher."""

from __future__ import annotations

from pathlib import Path

import pytest

from dlg import __main__ as launcher


@pytest.fixture
def calls(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> dict:
    """Replace uvicorn.run and threading.Timer with recorders."""
    rec: dict = {"timers": [], "config": str(tmp_path / "cfg.toml")}
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
    launcher.main(["--config", calls["config"]])
    args, kwargs = calls["run"]
    assert kwargs == {"host": "127.0.0.1", "port": 8000}
    assert callable(args[0])  # the FastAPI app object
    assert calls["timers"][0][2] == ("http://localhost:8000/",)


def test_no_browser_and_custom_port(calls: dict) -> None:
    """``--no-browser`` skips the browser; ``--port`` is honoured."""
    launcher.main(["--config", calls["config"], "--no-browser", "--port", "9001"])
    assert calls["timers"] == []
    assert calls["run"][1]["port"] == 9001


def test_config_created_and_port_from_config(calls: dict) -> None:
    """The config file's port is used unless --port is given."""
    cfg = Path(calls["config"])
    cfg.write_text("port = 9100\n", encoding="utf-8")
    launcher.main(["--config", str(cfg), "--no-browser"])
    assert calls["run"][1]["port"] == 9100


def test_missing_config_is_created(calls: dict) -> None:
    """A missing config file is created with defaults on start-up."""
    launcher.main(["--config", calls["config"], "--no-browser"])
    assert Path(calls["config"]).is_file()


def test_invalid_config_exits_with_message(calls: dict, capsys: pytest.CaptureFixture[str]) -> None:
    """A broken config file stops start-up with exit code 2 and a clear message."""
    Path(calls["config"]).write_text("port = [", encoding="utf-8")
    with pytest.raises(SystemExit) as exc:
        launcher.main(["--config", calls["config"]])
    assert exc.value.code == 2
    assert "not valid TOML" in capsys.readouterr().err
