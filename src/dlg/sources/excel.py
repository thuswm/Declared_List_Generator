"""Excel (.xlsx) implementation of :class:`~dlg.sources.base.ItemSource`.

Header conventions (first row of the first sheet):

* Every column with a non-empty header is shown and exported, in sheet order.
* A header starting with ``#`` (e.g. ``#Internal note``) marks a column the
  tool ignores completely.
* One column, named by the ``id_column`` setting, holds unique item IDs.
"""

from __future__ import annotations

from collections.abc import Iterator
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import load_workbook

from dlg.models import Column, Item, ItemTable

#: Default name of the column holding the unique item IDs.
DEFAULT_ID_COLUMN = "ID"

#: Header prefix marking a column that the tool ignores.
OMIT_MARKER = "#"


class ItemSourceError(Exception):
    """Raised when the item database is missing, malformed or inconsistent."""


def cell_text(value: object) -> str:
    """Convert an Excel cell value to display text.

    Dates without a time of day are shown as ``YYYY-MM-DD`` rather than with a
    ``00:00:00`` suffix, and whole-number floats lose the ``.0`` that Excel's
    number storage would otherwise add.

    Args:
        value: Raw value as returned by openpyxl.

    Returns:
        Stripped text ('' for empty cells).
    """
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat() if value.time() == time(0) else value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


class ExcelItemSource:
    """Reads items from the first sheet of an Excel workbook.

    Attributes:
        path: Location of the workbook.
        id_column: Header of the column holding the unique item IDs.
        name: Filename of the workbook.
    """

    def __init__(self, path: str | Path, id_column: str = DEFAULT_ID_COLUMN) -> None:
        """Initialise the source.

        Args:
            path: Path to the .xlsx file.
            id_column: Header of the ID column.
        """
        self.path = Path(path)
        self.id_column = id_column
        self.name = self.path.name

    def load(self) -> ItemTable:
        """Load the shown columns and all items from the workbook.

        Returns:
            Columns and items in sheet order.

        Raises:
            ItemSourceError: If the file is missing or unreadable, the ID column
                is absent, a header appears twice, or an ID appears twice.
        """
        if not self.path.is_file():
            raise ItemSourceError(f"Item database not found: {self.path}")
        try:
            wb = load_workbook(self.path, read_only=True, data_only=True)
        except Exception as exc:  # openpyxl raises various types on bad files
            raise ItemSourceError(f"Cannot read {self.path}: {exc}") from exc

        try:
            rows = wb.worksheets[0].iter_rows(values_only=True)
            header = next(rows, None)
            if header is None:
                raise ItemSourceError(f"{self.name} is empty (no header row).")
            positions = self._shown_columns(header)
            items = self._parse_rows(rows, positions)
        finally:
            wb.close()
        return ItemTable(
            id_column=self.id_column,
            columns=[Column(name=n) for n in positions],
            items=items,
        )

    def _shown_columns(self, header: tuple) -> dict[str, int]:
        """Work out which columns are shown and where they are.

        Args:
            header: The first row of the sheet.

        Returns:
            Column name to zero-based sheet index, in sheet order.

        Raises:
            ItemSourceError: If a header is duplicated or the ID column is missing.
        """
        positions: dict[str, int] = {}
        for i, raw in enumerate(header):
            name = cell_text(raw)
            if not name or name.startswith(OMIT_MARKER):
                continue
            if name in positions:
                raise ItemSourceError(f"Column '{name}' appears more than once in {self.name}.")
            positions[name] = i
        if self.id_column not in positions:
            shown = ", ".join(positions) or "(none)"
            raise ItemSourceError(
                f"ID column '{self.id_column}' not found in {self.name}. "
                f"Set 'id_column' in your configuration file to one of: {shown}."
            )
        return positions

    def _parse_rows(self, rows: Iterator[tuple], positions: dict[str, int]) -> list[Item]:
        """Convert sheet rows to Items.

        Args:
            rows: Iterator of data rows (header already consumed).
            positions: Shown column name to sheet index.

        Returns:
            Parsed items; rows with an empty ID are skipped.

        Raises:
            ItemSourceError: On duplicate IDs.
        """
        items: list[Item] = []
        seen: set[str] = set()
        for row in rows:
            # Rows can be shorter than the header when trailing cells are empty.
            values = {n: cell_text(row[i]) if i < len(row) else "" for n, i in positions.items()}
            item_id = values[self.id_column]
            if not item_id:
                continue  # blank/spacer row
            if item_id in seen:
                raise ItemSourceError(f"Duplicate item ID '{item_id}' in {self.name}.")
            seen.add(item_id)
            items.append(Item(id=item_id, values=values))
        return items
