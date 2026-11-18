"""Abstract interface for item databases.

Keeping this a ``Protocol`` lets the Excel source be swapped for SQLite or
another backend later without touching the rest of the application.
"""

from __future__ import annotations

from typing import Protocol

from dlg.models import ItemTable


class ItemSource(Protocol):
    """Anything that can provide the list of selectable items."""

    #: Short name of the underlying database (e.g. the filename).
    name: str

    def load(self) -> ItemTable:
        """Load the shown columns and all selectable items.

        Returns:
            Columns and items, in database order.
        """
        ...
