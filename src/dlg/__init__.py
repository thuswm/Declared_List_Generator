"""Declared List Generator: pick items in the browser, export a Word declared list."""

from __future__ import annotations

import tomllib
from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as _installed_version
from pathlib import Path

#: ``pyproject.toml`` of a source checkout (``src/dlg/__init__.py`` → project root).
_PYPROJECT = Path(__file__).resolve().parent.parent.parent / "pyproject.toml"


def get_version() -> str:
    """Return the application version.

    ``pyproject.toml`` is the single source of the version. With the editable
    install the README describes, the installed metadata keeps the version from
    the time of ``pip install`` and goes stale after a ``git pull``, so the
    checkout's ``pyproject.toml`` is read first. The installed metadata is the
    fallback for a regular (non-editable) install.

    Returns:
        The version string, or ``"unknown"`` if neither source is available.
    """
    try:
        with _PYPROJECT.open("rb") as fh:
            project = tomllib.load(fh).get("project", {})
        if project.get("name") == "dlg" and isinstance(project.get("version"), str):
            return project["version"]
    except (OSError, tomllib.TOMLDecodeError):
        pass
    try:
        return _installed_version("dlg")
    except PackageNotFoundError:
        return "unknown"
