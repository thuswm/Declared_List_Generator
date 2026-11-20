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
        sessions_dir: Folder for session JSON files (the user's library).
        global_sessions_dir: Shared session library (e.g. on a network drive);
            None when not set. Written as ``""`` in the file.
        table_layout_file: Workbook with the Word table layout sheet; None
            (``""``) means the item database workbook.
        table_layout_sheet: Name of the Word table layout sheet.
        table_font_size: Font size (pt) of the Word table; 0 keeps the
            template's font size.
        mark_empty_inputs: Write ``empty_input_text`` into empty cells of
            input (``*``) columns in the Word table.
        empty_input_text: Text for those cells (e.g. ``TBD``).
        output_dir: Folder where generated Word documents are written.
        id_column: Header of the item-database column holding the unique item IDs.
        host: Interface the server binds to (localhost: no authentication).
        port: Port the server listens on.
    """

    items_file: Path = field(default_factory=lambda: _default_path("data/items.xlsx"))
    template_file: Path = field(default_factory=lambda: _default_path("templates/declared_list_template.docx"))
    sessions_dir: Path = field(default_factory=lambda: _default_path("sessions"))
    global_sessions_dir: Path | None = None
    table_layout_file: Path | None = None
    table_layout_sheet: str = "Word layout"
    table_font_size: float = 0.0
    mark_empty_inputs: bool = False
    empty_input_text: str = "TBD"
    output_dir: Path = field(default_factory=lambda: _default_path("output"))
    id_column: str = "ID"
    host: str = "127.0.0.1"
    port: int = 8000


#: Smallest and largest table font size (pt); Word allows half-point steps.
MIN_FONT_SIZE = 4.0
MAX_FONT_SIZE = 72.0


def font_size_problem(size: float) -> str | None:
    """Check a table font size.

    Args:
        size: Size in points; 0 means "keep the template's size".

    Returns:
        A message describing the problem, or None if the size is valid.
    """
    if size == 0:
        return None
    if not MIN_FONT_SIZE <= size <= MAX_FONT_SIZE or (size * 2) != int(size * 2):
        return (f"Table font size must be 0 (template font) or {MIN_FONT_SIZE:g}–{MAX_FONT_SIZE:g} pt "
                "in steps of 0.5.")
    return None


#: Longest text accepted for empty input cells.
MAX_EMPTY_INPUT_TEXT = 100


def empty_input_text_problem(text: str) -> str | None:
    """Check the text written into empty input cells.

    Args:
        text: The text (already stripped).

    Returns:
        A message describing the problem, or None if the text is valid.
    """
    if not text:
        return "Text for empty input cells: enter a text such as TBD."
    if len(text) > MAX_EMPTY_INPUT_TEXT:
        return f"Text for empty input cells: at most {MAX_EMPTY_INPUT_TEXT} characters."
    return None


_PATH_KEYS = (
    "items_file", "template_file", "sessions_dir", "global_sessions_dir", "output_dir", "table_layout_file",
)
#: Path settings that may be unset (``""`` in the file, None in memory).
_OPTIONAL_PATH_KEYS = ("global_sessions_dir", "table_layout_file")


def _file_value(value: object) -> str | int | float | bool:
    """Convert a setting to what is written in the file (paths as text, None as ``""``).

    Args:
        value: Setting value.

    Returns:
        A string or number for :func:`_toml_value`.
    """
    if value is None:
        return ""
    return str(value) if isinstance(value, Path) else value  # type: ignore[return-value]


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


def _toml_value(value: str | int | float | bool) -> str:
    """Format a string, number or boolean as a TOML value.

    JSON string escaping is a valid subset of TOML basic strings, which keeps
    Windows backslashes and non-ASCII characters safe; JSON ``true``/``false``
    are TOML booleans too.

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
        lines.append(f"{f.name} = {_toml_value(_file_value(getattr(config, f.name)))}")
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

    _append_missing_keys(path, data)

    config = Config()
    for f in fields(config):
        if f.name not in data:
            continue
        value = data[f.name]
        if f.name == "table_font_size":
            # TOML writes 8 and 8.5 as int and float; both mean points.
            if not isinstance(value, (int, float)) or isinstance(value, bool):
                raise ConfigError(f"Configuration file {path}: 'table_font_size' must be a number.")
            value = float(value)
            problem = font_size_problem(value)
            if problem:
                raise ConfigError(f"Configuration file {path}: {problem}")
            setattr(config, f.name, value)
            continue
        if f.name == "mark_empty_inputs":
            if not isinstance(value, bool):
                raise ConfigError(f"Configuration file {path}: 'mark_empty_inputs' must be true or false.")
            config.mark_empty_inputs = value
            continue
        expected = int if f.name == "port" else str
        if not isinstance(value, expected) or isinstance(value, bool):
            raise ConfigError(f"Configuration file {path}: '{f.name}' must be a {expected.__name__}.")
        if f.name in _OPTIONAL_PATH_KEYS and not value.strip():
            value = None
        elif f.name in _PATH_KEYS:
            value = Path(value).expanduser()
            if not value.is_absolute():
                value = path.parent / value
        setattr(config, f.name, value)
    if not 1 <= config.port <= 65535:
        raise ConfigError(f"Configuration file {path}: 'port' must be between 1 and 65535.")
    if not config.id_column.strip():
        raise ConfigError(f"Configuration file {path}: 'id_column' must not be empty.")
    config.id_column = config.id_column.strip()
    if not config.table_layout_sheet.strip():
        raise ConfigError(f"Configuration file {path}: 'table_layout_sheet' must not be empty.")
    config.empty_input_text = config.empty_input_text.strip()
    problem = empty_input_text_problem(config.empty_input_text)
    if problem:
        raise ConfigError(f"Configuration file {path}: {problem}")
    return config


def _append_missing_keys(path: Path, data: dict[str, Any]) -> None:
    """Append settings added in newer versions to an existing file, with defaults.

    Appending (instead of rewriting) keeps the user's own edits and comments,
    while making new settings visible in the file so they can be changed.
    Failure to write is ignored: defaults still apply in memory.

    Args:
        path: Configuration file that was just read.
        data: Its parsed contents.
    """
    defaults = Config()
    missing = [f for f in fields(defaults) if f.name not in data]
    if not missing:
        return
    lines = ["", "# Added by a newer version of the tool (defaults):"]
    for f in missing:
        lines.append(f"{f.name} = {_toml_value(_file_value(getattr(defaults, f.name)))}")
    try:
        text = path.read_text(encoding="utf-8")
        sep = "" if text.endswith("\n") or not text else "\n"
        path.write_text(text + sep + "\n".join(lines) + "\n", encoding="utf-8")
    except OSError:
        pass
