"""Tests for the folder listing used by the browse dialog."""

from __future__ import annotations

import os
from pathlib import Path

import pytest

from dlg.filebrowser import BrowseError, BrowseNotFoundError, list_folder, roots


@pytest.fixture
def tree(tmp_path: Path) -> Path:
    """A folder with sub-folders, matching/non-matching and hidden files."""
    (tmp_path / "beta").mkdir()
    (tmp_path / "Alpha").mkdir()
    (tmp_path / ".hidden_dir").mkdir()
    for name in ("b.xlsx", "A.XLSX", "notes.txt", "~$b.xlsx", ".secret.xlsx", "c.xlsm"):
        (tmp_path / name).write_text("x")
    return tmp_path


def test_lists_folders_and_xlsx_files_sorted(tree: Path) -> None:
    """Folders and .xlsx files (any case) are listed, sorted case-insensitively; hidden/lock files are not."""
    listing = list_folder(tree)
    assert listing.path == str(tree.resolve())
    assert [d.name for d in listing.dirs] == ["Alpha", "beta"]
    assert [f.name for f in listing.files] == ["A.XLSX", "b.xlsx"]
    assert listing.files[1].path == str(tree.resolve() / "b.xlsx")
    assert listing.parent == str(tree.resolve().parent)


def test_folders_only(tree: Path) -> None:
    """With no extensions, only folders are listed (choosing an output folder)."""
    listing = list_folder(tree, ())
    assert listing.files == []
    assert len(listing.dirs) == 2


def test_file_path_lists_its_folder(tree: Path) -> None:
    """Opening at a file shows the folder containing it."""
    assert list_folder(tree / "b.xlsx").path == str(tree.resolve())


def test_empty_path_is_home() -> None:
    """No path means the home folder."""
    assert list_folder("").path == str(Path.home().resolve())


def test_root_has_no_parent() -> None:
    """At a file-system root there is no parent folder."""
    root = Path(Path.home().anchor)
    assert list_folder(root, ()).parent is None


def test_missing_folder(tmp_path: Path) -> None:
    """A missing folder raises BrowseNotFoundError."""
    with pytest.raises(BrowseNotFoundError):
        list_folder(tmp_path / "nope")


@pytest.mark.skipif(os.name == "nt" or os.geteuid() == 0, reason="POSIX permissions; root can read anything")
def test_unreadable_folder(tmp_path: Path) -> None:
    """A folder without read permission raises BrowseError."""
    locked = tmp_path / "locked"
    locked.mkdir()
    locked.chmod(0)
    try:
        with pytest.raises(BrowseError, match="Cannot open"):
            list_folder(locked)
    finally:
        locked.chmod(0o755)


def test_roots_per_platform() -> None:
    """Home is always offered; '/' on Linux/macOS."""
    linux = roots("linux")
    assert [r.name for r in linux] == ["Home", "/"]
    assert linux[0].path == str(Path.home())
    assert roots("win32")[0].name == "Home"
