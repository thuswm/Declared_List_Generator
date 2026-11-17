"""Pydantic data models shared across the Declared List Generator.

Defines the selectable ``Item`` read from the item database and the
``Session`` that stores a user's selection so it can be reloaded later.
"""

from __future__ import annotations

from datetime import datetime, timezone

from pydantic import BaseModel, Field

#: Current session file schema version. Bump when the format changes.
SESSION_SCHEMA_VERSION = 1


def _now() -> datetime:
    """Return the current time as a timezone-aware UTC datetime.

    Returns:
        The current UTC datetime.
    """
    return datetime.now(timezone.utc)


class Item(BaseModel):
    """A selectable item from the item database.

    Attributes:
        id: Unique identifier of the item within the database.
        name: Short display name.
        description: Longer free-text description.
        category: Grouping label used for filtering.
        extra: Any additional spreadsheet columns, keyed by header name.
    """

    id: str
    name: str
    description: str = ""
    category: str = ""
    extra: dict[str, str] = Field(default_factory=dict)


class Session(BaseModel):
    """A saved selection of items (a declared list configuration).

    Attributes:
        schema_version: Version of the session file format.
        name: Human-readable session name.
        created: Creation timestamp (UTC).
        modified: Last modification timestamp (UTC).
        source: Filename of the item database the selection was made from.
        selected_ids: IDs of the selected items, in selection order.
    """

    schema_version: int = SESSION_SCHEMA_VERSION
    name: str
    created: datetime = Field(default_factory=_now)
    modified: datetime = Field(default_factory=_now)
    source: str = ""
    selected_ids: list[str] = Field(default_factory=list)
