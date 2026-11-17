"""Tests for session persistence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from dlg.models import Session
from dlg.sessions import SessionError, SessionNotFoundError, SessionStore, safe_filename


def test_roundtrip(tmp_path: Path) -> None:
    """A saved session loads back identically (and the file is readable JSON)."""
    store = SessionStore(tmp_path / "s")  # directory is created on demand
    path = store.save(Session(name="Alpha", source="items.xlsx", selected_ids=["B", "A"]))
    assert json.loads(path.read_text())["selected_ids"] == ["B", "A"]
    loaded = store.load("Alpha")
    assert loaded.selected_ids == ["B", "A"]
    assert loaded.source == "items.xlsx"


def test_overwrite_keeps_created_and_updates_modified(tmp_path: Path) -> None:
    """Re-saving preserves ``created`` but advances ``modified``."""
    store = SessionStore(tmp_path)
    first = Session(name="x")
    store.save(first)
    created, modified = store.load("x").created, store.load("x").modified
    store.save(Session(name="x", selected_ids=["1"]))
    again = store.load("x")
    assert again.created == created
    assert again.modified >= modified
    assert again.selected_ids == ["1"]


def test_duplicate_ids_removed_in_order(tmp_path: Path) -> None:
    """Duplicate selected IDs are collapsed, keeping first occurrence order."""
    store = SessionStore(tmp_path)
    store.save(Session(name="d", selected_ids=["b", "a", "b"]))
    assert store.load("d").selected_ids == ["b", "a"]


@pytest.mark.parametrize("bad", ["", "   ", "..", "..."])
def test_invalid_names_rejected(bad: str, tmp_path: Path) -> None:
    """Names that sanitise to nothing are rejected."""
    with pytest.raises(SessionError):
        SessionStore(tmp_path).save(Session(name=bad))


def test_separators_are_replaced() -> None:
    """Path separators become underscores in the filename."""
    assert safe_filename("a/b\\c") == "a_b_c"


def test_path_traversal_cannot_escape(tmp_path: Path) -> None:
    """Separators in a name are neutralised; the file stays in the directory."""
    store = SessionStore(tmp_path / "sessions")
    path = store.save(Session(name="../../evil"))
    assert path.parent == tmp_path / "sessions"
    assert store.load("../../evil").name == "../../evil"  # display name preserved


def test_load_missing(tmp_path: Path) -> None:
    """Loading or deleting an unknown session raises SessionNotFoundError."""
    store = SessionStore(tmp_path)
    with pytest.raises(SessionNotFoundError):
        store.load("nope")
    with pytest.raises(SessionNotFoundError):
        store.delete("nope")


def test_corrupt_and_invalid_files(tmp_path: Path) -> None:
    """Bad JSON and schema violations raise SessionError on load."""
    store = SessionStore(tmp_path)
    (tmp_path / "bad.json").write_text("{not json")
    (tmp_path / "wrong.json").write_text(json.dumps({"schema_version": 1}))  # no name
    with pytest.raises(SessionError, match="Cannot read"):
        store.load("bad")
    with pytest.raises(SessionError, match="Invalid"):
        store.load("wrong")


def test_newer_schema_version_rejected(tmp_path: Path) -> None:
    """A file from a newer app version is refused with a clear message."""
    (tmp_path / "future.json").write_text(json.dumps({"schema_version": 99, "name": "f"}))
    with pytest.raises(SessionError, match="schema version 99"):
        SessionStore(tmp_path).load("future")


def test_list_sorted_and_skips_bad_files(tmp_path: Path) -> None:
    """Listing returns newest first and ignores unreadable files."""
    store = SessionStore(tmp_path)
    store.save(Session(name="old"))
    store.save(Session(name="new"))
    (tmp_path / "junk.json").write_text("garbage")
    assert [s.name for s in store.list()] == ["new", "old"]
    assert SessionStore(tmp_path / "missing").list() == []


def test_delete(tmp_path: Path) -> None:
    """Deleting removes the session."""
    store = SessionStore(tmp_path)
    store.save(Session(name="gone"))
    store.delete("gone")
    assert store.list() == []
