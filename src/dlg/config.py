"""Per-user configuration file (TOML).

Each user has a personal ``config.toml`` in the platform's standard
configuration folder. It is created with defaults on first use or when it is
missing. Reading uses the standard library (``tomllib``); the tiny writer below
only needs to handle flat string/integer settings.
"""

from __future__ import annotations

import json
import os
import sys
import tomllib
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parents[2]
APP_DIR_NAME = "dlg"
CONFIG_FILE_NAME = "config.toml"

_HEADER = """\
# Declared List Generator - personal settings.
# Relative paths are resolved against the folder containing this file.
# Delete this file to get the defaults back.
"""


class ConfigError(Exception):
    """Raised when the configuration file cannot be read or is invalid."""


def _default_path(name: str) -> Path:
    """Build the default location of a project-relative file or folder.

    Args:
        name: Path relative to the project root.

    Returns:
        Absolute path below the project root.
    """
    return PROJECT_ROOT / name


@dataclass
class Config:
    """Application settings.

    Attributes:
        items_file: Excel item database.
        template_file: Word template containing the placeholder.
        sessions_dir: Folder for session JSON files.
        output_dir: Folder where generated Word documents are written.
        host: Interface the server binds to (localhost: no authentication).
        port: Port the server listens on.
    """

    items_file: Path = field(default_factory=lambda: _default_path("data/items.xlsx"))
    template_file: Path = field(default_factory=lambda: _default_path("templates/declared_list_template.docx"))
    sessions_dir: Path = field(default_factory=lambda: _default_path("sessions"))
    output_dir: Path = field(default_factory=lambda: _default_path("output"))
    host: str = "127.0.0.1"
    port: int = 8000


_PATH_KEYS = ("items_file", "template_file", "sessions_dir", "output_dir")


def default_config_path(platform: str | None = None, env: dict[str, str] | None = None) -> Path:
    """Return the per-user configuration file location for this platform.

    Windows: ``%APPDATA%\\dlg\\config.toml``. Elsewhere (Linux, macOS):
    ``$XDG_CONFIG_HOME/dlg/config.toml``, falling back to ``~/.config/dlg/``.

    Args:
        platform: Value of ``sys.platform`` (injectable for tests).
        env: Environment mapping (injectable for tests).

    Returns:
        Absolute path of the configuration file (which may not exist yet).
    """
    platform = platform if platform is not None else sys.platform
    env = env if env is not None else dict(os.environ)
    if platform.startswith("win"):
        base = Path(env["APPDATA"]) if env.get("APPDATA") else Path.home() / "AppData" / "Roaming"
    else:
        base = Path(env["XDG_CONFIG_HOME"]) if env.get("XDG_CONFIG_HOME") else Path.home() / ".config"
    return base / APP_DIR_NAME / CONFIG_FILE_NAME


def _toml_value(value: str | int) -> str:
    """Format a string or integer as a TOML value.

    JSON string escaping is a valid subset of TOML basic strings, which keeps
    Windows backslashes and non-ASCII characters safe.

    Args:
        value: The value to format.

    Returns:
        TOML source text for the value.
    """
    return json.dumps(value, ensure_ascii=False)


def save_config(config: Config, path: Path) -> None:
    """Write the configuration file (atomically), creating folders as needed.

    Args:
        config: Settings to store.
        path: Target file.

    Raises:
        ConfigError: If the file cannot be written.
    """
    lines = [_HEADER]
    for f in fields(config):
        value = getattr(config, f.name)
        lines.append(f"{f.name} = {_toml_value(str(value) if isinstance(value, Path) else value)}")
    tmp = path.with_name(path.name + ".tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text("\n".join(lines) + "\n", encoding="utf-8")
        tmp.replace(path)  # atomic: never leaves a half-written config behind
    except OSError as exc:
        raise ConfigError(f"Cannot write configuration file {path}: {exc}") from exc


def load_config(path: Path | None = None) -> Config:
    """Load the configuration, creating the file with defaults if missing.

    Unknown keys are ignored so a newer/older file never blocks start-up.
    Keys that are absent keep their default.

    Args:
        path: Configuration file; defaults to :func:`default_config_path`.

    Returns:
        The settings, with paths made absolute.

    Raises:
        ConfigError: If the file is not valid TOML, a value has the wrong
            type, or the file cannot be read or created.
    """
    path = (path or default_config_path()).expanduser()
    if not path.exists():
        save_config(Config(), path)
    try:
        data: dict[str, Any] = tomllib.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise ConfigError(f"Cannot read configuration file {path}: {exc}") from exc
    except tomllib.TOMLDecodeError as exc:
        raise ConfigError(f"Configuration file {path} is not valid TOML: {exc}") from exc

    config = Config()
    for f in fields(config):
        if f.name not in data:
            continue
        value = data[f.name]
        expected = int if f.name == "port" else str
        if not isinstance(value, expected) or isinstance(value, bool):
            raise ConfigError(f"Configuration file {path}: '{f.name}' must be a {expected.__name__}.")
        if f.name in _PATH_KEYS:
            value = Path(value).expanduser()
            if not value.is_absolute():
                value = path.parent / value
        setattr(config, f.name, value)
    if not 1 <= config.port <= 65535:
        raise ConfigError(f"Configuration file {path}: 'port' must be between 1 and 65535.")
    return config
