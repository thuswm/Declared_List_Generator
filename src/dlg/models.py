"""Pydantic data models shared across the Declared List Generator.

Defines the item database contents (``ItemTable`` of ``Column`` and ``Item``)
and the ``Session`` that stores a user's selection so it can be reloaded later.
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


class Column(BaseModel):
    """A column of the item database that the tool shows and exports.

    Attributes:
        name: Header text from the database.
    """

    name: str


class Item(BaseModel):
    """A selectable item (one row of the item database).

    Attributes:
        id: Unique identifier of the item (value of the configured ID column).
        values: Cell text per column name, for every shown column, in column
            order (empty string for empty cells). Includes the ID column.
    """

    id: str
    values: dict[str, str] = Field(default_factory=dict)


class ItemTable(BaseModel):
    """The contents of the item database as shown in the tool.

    Attributes:
        id_column: Name of the column holding the unique item IDs.
        columns: Shown columns, in database order (omitted columns excluded).
        items: All items, in database order.
    """

    id_column: str
    columns: list[Column]
    items: list[Item]

    @property
    def column_names(self) -> list[str]:
        """Names of the shown columns, in order."""
        return [c.name for c in self.columns]


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
