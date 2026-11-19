"""Folder listing for the GUI's file/folder browser dialog.

Browsers cannot reveal real file paths, so the (local) server lists folders
for the GUI instead. Works the same on Linux and Windows using only the
standard library.
"""

from __future__ import annotations

import os
import stat
import string
import sys
from dataclasses import dataclass, field
from pathlib import Path


class BrowseError(Exception):
    """Raised when a folder cannot be listed."""


class BrowseNotFoundError(BrowseError):
    """Raised when the requested folder does not exist."""


@dataclass
class Entry:
    """One folder or file in a listing.

    Attributes:
        name: Display name.
        path: Absolute path.
    """

    name: str
    path: str


@dataclass
class Listing:
    """Contents of one folder, ready for the browser dialog.

    Attributes:
        path: Absolute path of the listed folder.
        parent: Its parent folder, or None at a file-system root.
        dirs: Sub-folders, sorted by name (case-insensitive).
        files: Files with an accepted extension, sorted by name.
        roots: Starting points (home folder, drives or ``/``).
    """

    path: str
    parent: str | None
    dirs: list[Entry] = field(default_factory=list)
    files: list[Entry] = field(default_factory=list)
    roots: list[Entry] = field(default_factory=list)


def roots(platform: str | None = None) -> list[Entry]:
    """Return the places the dialog offers as starting points.

    Args:
        platform: Value of ``sys.platform`` (injectable for tests).

    Returns:
        The home folder, then the drives (Windows) or ``/`` (elsewhere).
    """
    platform = platform if platform is not None else sys.platform
    result = [Entry(name="Home", path=str(Path.home()))]
    if platform.startswith("win"):
        if hasattr(os, "listdrives"):  # Python 3.12+
            drives = os.listdrives()
        else:  # pragma: no cover - Python 3.11 on Windows
            drives = [f"{c}:\\" for c in string.ascii_uppercase if os.path.exists(f"{c}:\\")]
        result += [Entry(name=d.rstrip("\\"), path=d) for d in drives]
    else:
        result.append(Entry(name="/", path="/"))
    return result


def _is_hidden(entry: os.DirEntry[str]) -> bool:
    """Decide whether to hide an entry from the dialog.

    Dot-files (Unix convention), Office lock files (``~$name.xlsx``/``~$name.docx``,
    which Excel and Word leave next to open files) and files with the Windows hidden
    attribute are hidden.

    Args:
        entry: Directory entry.

    Returns:
        True if the entry should not be listed.
    """
    if entry.name.startswith((".", "~$")):
        return True
    try:
        attributes = getattr(entry.stat(follow_symlinks=False), "st_file_attributes", 0)
    except OSError:
        return True
    return bool(attributes & stat.FILE_ATTRIBUTE_HIDDEN)


def list_folder(path: str | Path | None, extensions: tuple[str, ...] = (".xlsx",)) -> Listing:
    """List the sub-folders and matching files of a folder.

    Args:
        path: Folder to list. Empty/None means the home folder; a file means
            the folder containing it (so the dialog opens at the current file).
        extensions: Lower-case file extensions to list; empty lists no files.

    Returns:
        The listing.

    Raises:
        BrowseNotFoundError: If the folder does not exist.
        BrowseError: If the folder cannot be read (e.g. no permission).
    """
    folder = Path(path).expanduser() if path else Path.home()
    if folder.is_file():
        folder = folder.parent
    if not folder.is_dir():
        raise BrowseNotFoundError(f"Folder not found: {folder}")
    folder = folder.resolve()

    listing = Listing(
        path=str(folder),
        parent=str(folder.parent) if folder.parent != folder else None,
        roots=roots(),
    )
    try:
        with os.scandir(folder) as it:
            for entry in it:
                if _is_hidden(entry):
                    continue
                try:
                    is_dir = entry.is_dir()
                except OSError:
                    continue  # broken link or vanished entry
                if is_dir:
                    listing.dirs.append(Entry(entry.name, entry.path))
                elif extensions and entry.name.lower().endswith(extensions):
                    listing.files.append(Entry(entry.name, entry.path))
    except OSError as exc:  # PermissionError and friends
        raise BrowseError(f"Cannot open folder {folder}: {exc.strerror or exc}") from exc

    def key(e: Entry) -> str:
        return e.name.casefold()

    listing.dirs.sort(key=key)
    listing.files.sort(key=key)
    return listing
