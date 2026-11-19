"""Tests for the per-user configuration file."""

from __future__ import annotations

from pathlib import Path

import pytest

from dlg.config import Config, ConfigError, default_config_path, load_config, save_config


def test_default_path_linux_uses_xdg_or_home() -> None:
    """Linux/macOS: $XDG_CONFIG_HOME wins, else ~/.config."""
    p = default_config_path("linux", {"XDG_CONFIG_HOME": "/x/cfg"})
    assert p == Path("/x/cfg/dlg/config.toml")
    assert default_config_path("linux", {}) == Path.home() / ".config" / "dlg" / "config.toml"


def test_default_path_windows_uses_appdata() -> None:
    """Windows: %APPDATA%\\dlg\\config.toml, with a home-based fallback."""
    p = default_config_path("win32", {"APPDATA": "/appdata"})
    assert p == Path("/appdata") / "dlg" / "config.toml"
    assert default_config_path("win32", {}) == Path.home() / "AppData" / "Roaming" / "dlg" / "config.toml"


def test_missing_file_is_created_with_defaults(tmp_path: Path) -> None:
    """First run: the file (and its folder) is created and defaults are returned."""
    path = tmp_path / "sub" / "config.toml"
    cfg = load_config(path)
    assert path.is_file()
    assert cfg == Config()
    assert load_config(path) == cfg  # round-trips


def test_values_and_relative_paths(tmp_path: Path) -> None:
    """Relative paths resolve against the config file's folder; absent keys keep defaults."""
    path = tmp_path / "config.toml"
    path.write_text('items_file = "db/items.xlsx"\noutput_dir = "out"\nport = 9000\nunknown = 1\n', encoding="utf-8")
    cfg = load_config(path)
    assert cfg.items_file == tmp_path / "db" / "items.xlsx"
    assert cfg.output_dir == tmp_path / "out"
    assert cfg.port == 9000
    assert cfg.template_file == Config().template_file


def test_windows_style_and_unicode_paths_round_trip(tmp_path: Path) -> None:
    """Backslashes and non-ASCII characters survive writing and reading."""
    path = tmp_path / "config.toml"
    cfg = Config(host="localhost")
    cfg.output_dir = Path("/data/Ünï \\ out")
    save_config(cfg, path)
    assert load_config(path).output_dir == cfg.output_dir


@pytest.mark.parametrize(
    "text, fragment",
    [("port = [", "not valid TOML"), ('port = "80"', "'port' must be a int"),
     ("items_file = 3", "'items_file' must be a str"), ("port = 70000", "between 1 and 65535"),
     ('id_column = "  "', "'id_column' must not be empty")],
)
def test_invalid_config(tmp_path: Path, text: str, fragment: str) -> None:
    """Broken files raise ConfigError with a helpful message."""
    path = tmp_path / "config.toml"
    path.write_text(text, encoding="utf-8")
    with pytest.raises(ConfigError, match=fragment):
        load_config(path)


def test_id_column_setting(tmp_path: Path) -> None:
    """``id_column`` defaults to 'ID' and is read (trimmed) from the file."""
    assert Config().id_column == "ID"
    path = tmp_path / "config.toml"
    path.write_text('id_column = " Part No "\n', encoding="utf-8")
    assert load_config(path).id_column == "Part No"


def test_new_settings_are_appended_to_an_old_file(tmp_path: Path) -> None:
    """Keys missing from an existing file are appended with defaults; user lines are kept."""
    path = tmp_path / "config.toml"
    path.write_text('# my comment\nport = 9000', encoding="utf-8")  # no trailing newline
    cfg = load_config(path)
    text = path.read_text(encoding="utf-8")
    assert text.startswith("# my comment\nport = 9000\n")
    assert 'id_column = "ID"' in text
    assert text.count("port =") == 1
    assert cfg.port == 9000
    assert load_config(path) == cfg
    assert path.read_text(encoding="utf-8") == text  # nothing appended twice


def test_global_sessions_dir_is_optional(tmp_path: Path) -> None:
    """The shared library is unset by default (written as ""), and a set value round-trips."""
    path = tmp_path / "config.toml"
    cfg = load_config(path)
    assert cfg.global_sessions_dir is None
    assert 'global_sessions_dir = ""' in path.read_text(encoding="utf-8")
    cfg.global_sessions_dir = tmp_path / "shared lib"
    save_config(cfg, path)
    assert load_config(path).global_sessions_dir == tmp_path / "shared lib"


def test_old_file_gets_global_sessions_dir_appended_unset(tmp_path: Path) -> None:
    """A 0.2.0 config file gains the new key, unset."""
    path = tmp_path / "config.toml"
    path.write_text('port = 9000\n', encoding="utf-8")
    assert load_config(path).global_sessions_dir is None
    assert 'global_sessions_dir = ""' in path.read_text(encoding="utf-8")
