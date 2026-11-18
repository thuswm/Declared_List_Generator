"""Persistence of sessions as human-readable JSON files.

Each session is stored as ``<directory>/<file>.json``. The session is
identified by the exact ``name`` stored *inside* the file; the filename is
only a readable, sanitised form of it. Because different names can sanitise
to the same filename (``a/b`` and ``a_b``, or ``Proj A`` and ``proj a`` on
case-insensitive Windows file systems), a free filename is chosen on first
save (``a_b.json``, ``a_b (2).json``, ...) and lookups match on the stored name.
"""

from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from pydantic import ValidationError

from dlg.models import SESSION_SCHEMA_VERSION, Session

#: File stems Windows reserves for devices (with any extension); never usable.
WINDOWS_RESERVED_NAMES = frozenset(
    ["CON", "PRN", "AUX", "NUL"] + [f"{p}{n}" for p in ("COM", "LPT") for n in range(1, 10)]
)

#: Upper bound for numbered alternative filenames; far beyond realistic use.
_MAX_SUFFIX = 1000


class SessionError(Exception):
    """Raised when a session cannot be saved, found or parsed."""


class SessionNotFoundError(SessionError):
    """Raised when a requested session file does not exist."""


def safe_filename(name: str) -> str:
    """Convert a session name to a safe filename stem.

    Anything outside letters, digits, space, ``.``, ``_`` and ``-`` becomes
    ``_``. This blocks path separators, so names like ``../x`` cannot escape
    the sessions directory. Windows device names (``CON``, ``NUL``, ...) get a
    leading ``_`` so the file can be created on Windows too.

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
    if stem.split(".")[0].upper() in WINDOWS_RESERVED_NAMES:
        stem = f"_{stem}"
    return stem


def migrate(data: dict[str, Any]) -> dict[str, Any]:
    """Upgrade raw session data from older schema versions to the current one.

    Each step upgrades one version, so future changes add one more step.

    Args:
        data: Parsed JSON of a session file (not modified).

    Returns:
        Data in the current schema version.
    """
    data = dict(data)
    version = data.get("schema_version", 1)
    if isinstance(version, int) and version < 2:
        # v1 -> v2: new fields configuration_name and user_values; model defaults fill them.
        data["schema_version"] = 2
    return data


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

    def _preferred_path(self, name: str) -> Path:
        """Return the natural file path for a session name.

        Args:
            name: Session name.

        Returns:
            ``<directory>/<safe name>.json`` (may belong to another session).

        Raises:
            SessionError: If the name is invalid.
        """
        return self.directory / f"{safe_filename(name)}.json"

    def _find(self, name: str) -> Path | None:
        """Locate the file of the session with exactly this name.

        The preferred path is tried first (the common case). Only if it holds a
        different session is the folder scanned for a file with this name.

        Args:
            name: Session name.

        Returns:
            The file, or None if no session has this name.

        Raises:
            SessionError: If the name is invalid, or the preferred file exists
                but cannot be read (reported rather than silently bypassed).
        """
        preferred = self._preferred_path(name)
        if preferred.is_file():
            if self._read(preferred).name == name:
                return preferred
        if not self.directory.is_dir():
            return None
        for path in sorted(self.directory.glob("*.json")):
            if path == preferred:
                continue
            try:
                if self._read(path).name == name:
                    return path
            except SessionError:
                continue  # unrelated broken file
        return None

    def _free_path(self, name: str) -> Path:
        """Choose an unused file path for a new session.

        Args:
            name: Session name.

        Returns:
            ``<stem>.json``, or ``<stem> (2).json``, ``(3)``, ... if taken.

        Raises:
            SessionError: If the name is invalid or no free name is found.
        """
        stem = safe_filename(name)
        for n in range(1, _MAX_SUFFIX):
            # exists() follows the file system's rules, so case-insensitive
            # collisions on Windows are caught here too.
            path = self.directory / (f"{stem}.json" if n == 1 else f"{stem} ({n}).json")
            if not path.exists():
                return path
        raise SessionError(f"Too many sessions with names similar to {name!r}")

    def save(self, session: Session) -> Path:
        """Write a session, overwriting any session with exactly the same name.

        The original ``created`` timestamp is kept when overwriting, and
        ``modified`` is set to now. Duplicate IDs are removed, and only
        non-empty user values of selected items are kept. The write is atomic
        (temp file + rename) so a crash cannot leave a half-written session.

        Args:
            session: The session to store.

        Returns:
            Path of the written file.

        Raises:
            SessionError: If the name is invalid or the session's existing file
                is unreadable.
        """
        path = self._find(session.name)
        if path is not None:
            session.created = self._read(path).created
        else:
            path = self._free_path(session.name)
        session.schema_version = SESSION_SCHEMA_VERSION
        session.modified = datetime.now(timezone.utc)
        session.selected_ids = list(dict.fromkeys(session.selected_ids))  # dedupe, keep order
        session.user_values = self._clean_user_values(session)

        self.directory.mkdir(parents=True, exist_ok=True)
        tmp = path.with_name(path.name + ".tmp")
        tmp.write_text(session.model_dump_json(indent=2) + "\n", encoding="utf-8")
        tmp.replace(path)
        return path

    @staticmethod
    def _clean_user_values(session: Session) -> dict[str, dict[str, str]]:
        """Keep only non-empty values of selected items, in selection order.

        Args:
            session: The session being saved.

        Returns:
            The cleaned ``user_values``.
        """
        cleaned: dict[str, dict[str, str]] = {}
        for item_id in session.selected_ids:
            values = {k: v for k, v in session.user_values.get(item_id, {}).items() if v.strip()}
            if values:
                cleaned[item_id] = values
        return cleaned

    def load(self, name: str) -> Session:
        """Load a session by name.

        Args:
            name: Session name.

        Returns:
            The parsed session, migrated to the current schema version.

        Raises:
            SessionNotFoundError: If no such session exists.
            SessionError: If the file is not valid JSON, fails validation, or
                was written by a newer, unsupported schema version.
        """
        path = self._find(name)
        if path is None:
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
        path = self._find(name)
        if path is None:
            raise SessionNotFoundError(f"Session not found: {name}")
        path.unlink()

    @staticmethod
    def _read(path: Path) -> Session:
        """Parse, migrate and validate one session file.

        Args:
            path: File to read.

        Returns:
            The parsed session in the current schema version.

        Raises:
            SessionError: On bad JSON, failed validation or a too-new version.
        """
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            raise SessionError(f"Cannot read session file {path.name}: {exc}") from exc
        if not isinstance(data, dict):
            raise SessionError(f"Invalid session file {path.name}: not a JSON object")
        version = data.get("schema_version")
        if isinstance(version, int) and version > SESSION_SCHEMA_VERSION:
            raise SessionError(
                f"{path.name} uses schema version {version}; "
                f"this app supports up to {SESSION_SCHEMA_VERSION}"
            )
        try:
            return Session.model_validate(migrate(data))
        except ValidationError as exc:
            raise SessionError(f"Invalid session file {path.name}: {exc}") from exc
