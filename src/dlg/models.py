"""Pydantic data models shared across the Declared List Generator.

Defines the item database contents (``ItemTable`` of ``Column`` and ``Item``)
and the ``Session`` that stores a user's selection so it can be reloaded later.
"""

from __future__ import annotations

from datetime import datetime, timezone

from pydantic import BaseModel, Field

#: Current session file schema version. Bump when the format changes and add a
#: migration step in ``dlg.sessions``.
SESSION_SCHEMA_VERSION = 2


def _now() -> datetime:
    """Return the current time as a timezone-aware UTC datetime.

    Returns:
        The current UTC datetime.
    """
    return datetime.now(timezone.utc)


class Column(BaseModel):
    """A column of the item database that the tool shows and exports.

    Attributes:
        name: Header text from the database (markers removed).
        user_input: True if the user types this column's values per item in the
            tool (header marked with a trailing ``*`` in the database).
    """

    name: str
    user_input: bool = False


class Item(BaseModel):
    """A selectable item (one row of the item database).

    Attributes:
        id: Unique identifier of the item (value of the configured ID column).
        values: Cell text per column name, for every shown column, in column
            order (empty string for empty cells). Includes the ID column.
            User-input columns are empty here; see ``ItemTable.with_user_values``.
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

    def with_user_values(self, user_values: dict[str, dict[str, str]]) -> list[Item]:
        """Return the items with the user's values filled into the input columns.

        Values for columns that are not input columns are ignored, so a session
        can never overwrite database data.

        Args:
            user_values: Typed values per item ID and column name.

        Returns:
            New Item objects (the table is not modified), in database order.
        """
        inputs = [c.name for c in self.columns if c.user_input]
        if not inputs:
            return list(self.items)
        merged = []
        for item in self.items:
            typed = user_values.get(item.id, {})
            values = {**item.values, **{c: typed.get(c, "") for c in inputs}}
            merged.append(Item(id=item.id, values=values))
        return merged


class Session(BaseModel):
    """A saved selection of items (a declared list configuration).

    Attributes:
        schema_version: Version of the session file format.
        name: Human-readable session name (unique key of the session).
        configuration_name: Free text describing the configuration the
            selected items belong to (added in schema version 2).
        created: Creation timestamp (UTC).
        modified: Last modification timestamp (UTC).
        source: Filename of the item database the selection was made from.
        selected_ids: IDs of the selected items, in selection order.
        user_values: Values typed by the user, per item ID and column name
            (added in schema version 2).
    """

    schema_version: int = SESSION_SCHEMA_VERSION
    name: str
    configuration_name: str = ""
    created: datetime = Field(default_factory=_now)
    modified: datetime = Field(default_factory=_now)
    source: str = ""
    selected_ids: list[str] = Field(default_factory=list)
    user_values: dict[str, dict[str, str]] = Field(default_factory=dict)
