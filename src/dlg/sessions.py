"""Persistence of sessions as human-readable JSON files.

Each session is stored as ``<directory>/<safe-name>.json``. The display name
inside the file is preserved exactly; only the filename is sanitised.
"""

from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path

from pydantic import ValidationError

from dlg.models import SESSION_SCHEMA_VERSION, Session


class SessionError(Exception):
    """Raised when a session cannot be saved, found or parsed."""


class SessionNotFoundError(SessionError):
    """Raised when a requested session file does not exist."""


def safe_filename(name: str) -> str:
    """Convert a session name to a safe filename stem.

    Anything outside letters, digits, space, ``.``, ``_`` and ``-`` becomes
    ``_``. This blocks path separators, so names like ``../x`` cannot escape
    the sessions directory.

    Args:
        name: The user-supplied session name.

    Returns:
        A sanitised stem without extension.

    Raises:
        SessionError: If nothing usable remains (empty, or only dots/spaces).
    """
    stem = re.sub(r"[^\w .-]", "_", name, flags=re.ASCII).strip(" .")
    if not stem:
        raise SessionError(f"Invalid session name: {name!r}")
    return stem


class SessionStore:
    """Reads and writes session files in a directory.

    Attributes:
        directory: Folder holding the ``.json`` session files.
    """

    def __init__(self, directory: str | Path) -> None:
        """Initialise the store.

        Args:
            directory: Session folder; created on first save if missing.
        """
        self.directory = Path(directory)

    def _path(self, name: str) -> Path:
        """Return the file path for a session name.

        Args:
            name: Session name.

        Returns:
            Path of the session's JSON file.
        """
        return self.directory / f"{safe_filename(name)}.json"

    def save(self, session: Session) -> Path:
        """Write a session, overwriting any session of the same name.

        The original ``created`` timestamp is kept when overwriting, and
        ``modified`` is set to now. The write is atomic (temp file + rename)
        so a crash cannot leave a half-written session.

        Args:
            session: The session to store.

        Returns:
            Path of the written file.

        Raises:
            SessionError: If the name is invalid or an existing file at that
                path is unreadable.
        """
        path = self._path(session.name)
        if path.exists():
            session.created = self.load(session.name).created
        session.modified = datetime.now(timezone.utc)
        session.selected_ids = list(dict.fromkeys(session.selected_ids))  # dedupe, keep order

        self.directory.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".json.tmp")
        tmp.write_text(session.model_dump_json(indent=2) + "\n", encoding="utf-8")
        tmp.replace(path)
        return path

    def load(self, name: str) -> Session:
        """Load a session by name.

        Args:
            name: Session name.

        Returns:
            The parsed session.

        Raises:
            SessionNotFoundError: If no such session exists.
            SessionError: If the file is not valid JSON, fails validation, or
                was written by a newer, unsupported schema version.
        """
        path = self._path(name)
        if not path.is_file():
            raise SessionNotFoundError(f"Session not found: {name}")
        return self._read(path)

    def list(self) -> list[Session]:
        """List all readable sessions, most recently modified first.

        Unreadable or invalid files are skipped so one bad file does not hide
        the rest.

        Returns:
            Parsed sessions.
        """
        if not self.directory.is_dir():
            return []
        sessions: list[Session] = []
        for path in self.directory.glob("*.json"):
            try:
                sessions.append(self._read(path))
            except SessionError:
                continue
        return sorted(sessions, key=lambda s: s.modified, reverse=True)

    def delete(self, name: str) -> None:
        """Delete a session.

        Args:
            name: Session name.

        Raises:
            SessionNotFoundError: If no such session exists.
        """
        path = self._path(name)
        if not path.is_file():
            raise SessionNotFoundError(f"Session not found: {name}")
        path.unlink()

    @staticmethod
    def _read(path: Path) -> Session:
        """Parse and validate one session file.

        Args:
            path: File to read.

        Returns:
            The parsed session.

        Raises:
            SessionError: On bad JSON, failed validation or a too-new version.
        """
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise SessionError(f"Cannot read session file {path.name}: {exc}") from exc
        version = data.get("schema_version") if isinstance(data, dict) else None
        if isinstance(version, int) and version > SESSION_SCHEMA_VERSION:
            raise SessionError(
                f"{path.name} uses schema version {version}; "
                f"this app supports up to {SESSION_SCHEMA_VERSION}"
            )
        try:
            return Session.model_validate(data)
        except ValidationError as exc:
            raise SessionError(f"Invalid session file {path.name}: {exc}") from exc
