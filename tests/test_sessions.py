"""Tests for session persistence."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from dlg.models import SESSION_SCHEMA_VERSION, Session
from dlg.sessions import SessionError, SessionNotFoundError, SessionStore, migrate, safe_filename


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
    (tmp_path / "list.json").write_text("[1, 2]")
    with pytest.raises(SessionError, match="Cannot read"):
        store.load("bad")
    with pytest.raises(SessionError, match="Invalid"):
        store.load("wrong")
    with pytest.raises(SessionError, match="not a JSON object"):
        store.load("list")


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


V1_FILE = {
    "schema_version": 1,
    "name": "Legacy",
    "created": "2026-11-17T18:34:41Z",
    "modified": "2026-11-17T18:34:41Z",
    "source": "items.xlsx",
    "selected_ids": ["DL-001", "DL-003"],
}


def test_v2_fields_roundtrip(tmp_path: Path) -> None:
    """Configuration name and user values are stored and read back."""
    store = SessionStore(tmp_path)
    path = store.save(Session(
        name="A", configuration_name="Rig 7, build B", selected_ids=["1", "2"],
        user_values={"1": {"Serial number": "SN-1"}},
    ))
    raw = json.loads(path.read_text(encoding="utf-8"))
    assert raw["schema_version"] == SESSION_SCHEMA_VERSION == 2
    loaded = store.load("A")
    assert loaded.configuration_name == "Rig 7, build B"
    assert loaded.user_values == {"1": {"Serial number": "SN-1"}}


def test_user_values_only_for_selected_items_and_non_empty(tmp_path: Path) -> None:
    """Values of unselected items and blank values are not saved."""
    store = SessionStore(tmp_path)
    store.save(Session(
        name="u", selected_ids=["b", "a"],
        user_values={"a": {"S": "1", "T": "  "}, "b": {"S": ""}, "gone": {"S": "x"}},
    ))
    assert store.load("u").user_values == {"a": {"S": "1"}}


def test_v1_file_is_migrated(tmp_path: Path) -> None:
    """A schema v1 file loads with empty v2 fields and is written as v2 on save."""
    (tmp_path / "Legacy.json").write_text(json.dumps(V1_FILE), encoding="utf-8")
    store = SessionStore(tmp_path)
    s = store.load("Legacy")
    assert s.schema_version == 2
    assert s.configuration_name == "" and s.user_values == {}
    assert s.selected_ids == ["DL-001", "DL-003"]
    assert json.loads((tmp_path / "Legacy.json").read_text())["schema_version"] == 1  # untouched until saved
    s.configuration_name = "Now named"
    store.save(s)
    raw = json.loads((tmp_path / "Legacy.json").read_text())
    assert raw["schema_version"] == 2 and raw["configuration_name"] == "Now named"
    assert raw["created"].startswith("2026-11-17T18:34:41")


def test_migrate_does_not_modify_input() -> None:
    """``migrate`` returns a new dict; a missing version counts as v1."""
    data = {"name": "x"}
    assert migrate(data)["schema_version"] == 2
    assert data == {"name": "x"}


def test_colliding_names_get_separate_files(tmp_path: Path) -> None:
    """Names that sanitise alike ('a/b', 'a_b', 'a:b') no longer overwrite each other."""
    store = SessionStore(tmp_path)
    paths = [store.save(Session(name=n, selected_ids=[n])) for n in ("a/b", "a_b", "a:b")]
    assert [p.name for p in paths] == ["a_b.json", "a_b (2).json", "a_b (3).json"]
    for n in ("a/b", "a_b", "a:b"):
        assert store.load(n).selected_ids == [n]
    assert store.save(Session(name="a_b", selected_ids=["again"])).name == "a_b (2).json"  # overwrite in place
    store.delete("a/b")
    assert sorted(s.name for s in store.list()) == ["a:b", "a_b"]
    assert store.save(Session(name="a?b")).name == "a_b.json"  # freed name is reused


def test_name_differing_only_in_case(tmp_path: Path) -> None:
    """'Proj' and 'proj' are separate sessions, also where the file system ignores case."""
    store = SessionStore(tmp_path)
    store.save(Session(name="Proj", selected_ids=["1"]))
    store.save(Session(name="proj", selected_ids=["2"]))
    assert store.load("Proj").selected_ids == ["1"]
    assert store.load("proj").selected_ids == ["2"]
    assert len(list(tmp_path.glob("*.json"))) == 2


def test_v1_file_with_sanitised_name_is_found(tmp_path: Path) -> None:
    """Old files saved under a sanitised filename are still found by their real name."""
    (tmp_path / "a_b.json").write_text(json.dumps({**V1_FILE, "name": "a/b"}), encoding="utf-8")
    store = SessionStore(tmp_path)
    assert store.load("a/b").name == "a/b"
    with pytest.raises(SessionNotFoundError):
        store.load("a_b")
    assert store.save(Session(name="a_b")).name == "a_b (2).json"


@pytest.mark.parametrize("name, stem", [("CON", "_CON"), ("nul", "_nul"), ("com1.txt", "_com1.txt"), ("CONSOLE", "CONSOLE")])
def test_windows_reserved_names(name: str, stem: str) -> None:
    """Windows device names cannot be used as file names, so they are prefixed."""
    assert safe_filename(name) == stem
