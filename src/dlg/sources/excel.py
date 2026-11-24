"""Excel (.xlsx) implementation of :class:`~dlg.sources.base.ItemSource`.

Header conventions (first row of the data sheet, which is the first sheet
other than the Word layout sheet, see ``dlg.layout``):

* Every column with a non-empty header is shown and exported, in sheet order.
* A header starting with ``#`` (e.g. ``#Internal note``) marks a column the
  tool ignores completely.
* A header ending with ``*`` (e.g. ``Serial number*``) marks a user-input
  column: the user types its value per item in the tool. The ``*`` is not part
  of the column name, and the sheet's cells in that column are ignored.
* One column, named by the ``id_column`` setting, holds unique item IDs.
"""

from __future__ import annotations

from collections.abc import Iterator
from datetime import date, datetime, time
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from dlg.models import Column, Item, ItemTable

#: Default name of the column holding the unique item IDs.
DEFAULT_ID_COLUMN = "ID"

#: Header prefix marking a column that the tool ignores.
OMIT_MARKER = "#"

#: Header suffix marking a column whose values the user types in the tool.
INPUT_MARKER = "*"


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

    def __init__(
        self, path: str | Path, id_column: str = DEFAULT_ID_COLUMN, skip_sheet: str | None = None
    ) -> None:
        """Initialise the source.

        Args:
            path: Path to the .xlsx file.
            id_column: Header of the ID column.
            skip_sheet: Sheet that never holds the data (the Word layout sheet),
                matched ignoring case; None to always use the first sheet.
        """
        self.path = Path(path)
        self.id_column = id_column
        self.skip_sheet = skip_sheet
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
            rows = self._data_sheet(wb).iter_rows(values_only=True)
            header = next(rows, None)
            if header is None:
                raise ItemSourceError(f"{self.name} is empty (no header row).")
            columns, positions = self._shown_columns(header)
            items = self._parse_rows(rows, columns, positions)
        finally:
            wb.close()
        return ItemTable(id_column=self.id_column, columns=columns, items=items)

    def headers(self) -> list[str]:
        """Read the headers that can serve as the ID column (for the Settings pick list).

        Only the header row is read, and no ID column is needed, so this works
        before the ID column is set correctly.

        Returns:
            Shown headers in sheet order, without omitted (``#``) and
            user-input (``*``) columns, which cannot hold the IDs.

        Raises:
            ItemSourceError: If the file is missing or unreadable or has no data sheet.
        """
        if not self.path.is_file():
            raise ItemSourceError(f"Item database not found: {self.path}")
        try:
            wb = load_workbook(self.path, read_only=True, data_only=True)
        except Exception as exc:  # openpyxl raises various types on bad files
            raise ItemSourceError(f"Cannot read {self.path}: {exc}") from exc
        try:
            header = next(self._data_sheet(wb).iter_rows(values_only=True, max_row=1), ())
        finally:
            wb.close()
        names: list[str] = []
        for raw in header:
            name = cell_text(raw)
            if name and not name.startswith(OMIT_MARKER) and not name.endswith(INPUT_MARKER) and name not in names:
                names.append(name)
        return names

    def _data_sheet(self, wb: Workbook) -> Worksheet:
        """Pick the data sheet: the first sheet that is not the layout sheet.

        Args:
            wb: The open workbook.

        Returns:
            The worksheet holding the items.

        Raises:
            ItemSourceError: If the workbook has no other sheet.
        """
        skip = (self.skip_sheet or "").strip().casefold()
        for ws in wb.worksheets:
            if not skip or ws.title.strip().casefold() != skip:
                return ws
        raise ItemSourceError(f"{self.name} has no data sheet (only the layout sheet '{self.skip_sheet}').")

    def _shown_columns(self, header: tuple) -> tuple[list[Column], dict[str, int]]:
        """Work out which columns are shown, which are user input, and where they are.

        Args:
            header: The first row of the sheet.

        Returns:
            The shown columns in sheet order, and column name to zero-based
            sheet index.

        Raises:
            ItemSourceError: If a header is duplicated (after removing the
                ``*``), the ID column is missing, or the ID column is marked
                as user input.
        """
        columns: list[Column] = []
        positions: dict[str, int] = {}
        for i, raw in enumerate(header):
            name = cell_text(raw)
            if not name or name.startswith(OMIT_MARKER):
                continue
            user_input = name.endswith(INPUT_MARKER)
            if user_input:
                name = name.rstrip(INPUT_MARKER).rstrip()
                if not name:
                    continue  # a header of only "*" has no usable name
            if name in positions:
                raise ItemSourceError(f"Column '{name}' appears more than once in {self.name}.")
            positions[name] = i
            columns.append(Column(name=name, user_input=user_input))
        if any(c.user_input and c.name == self.id_column for c in columns):
            raise ItemSourceError(
                f"ID column '{self.id_column}' in {self.name} cannot be a user-input column (remove the '*')."
            )
        if self.id_column not in positions:
            shown = ", ".join(positions) or "(none)"
            raise ItemSourceError(
                f"ID column '{self.id_column}' not found in {self.name}. "
                f"Set the ID column in Settings ('id_column' in the configuration file) to one of: {shown}."
            )
        return columns, positions

    def _parse_rows(
        self, rows: Iterator[tuple], columns: list[Column], positions: dict[str, int]
    ) -> list[Item]:
        """Convert sheet rows to Items.

        Args:
            rows: Iterator of data rows (header already consumed).
            columns: Shown columns (user-input columns get empty values).
            positions: Shown column name to sheet index.

        Returns:
            Parsed items; rows with an empty ID are skipped.

        Raises:
            ItemSourceError: On duplicate IDs.
        """
        items: list[Item] = []
        seen: set[str] = set()
        inputs = {c.name for c in columns if c.user_input}
        for row in rows:
            # Rows can be shorter than the header when trailing cells are empty.
            values = {
                n: "" if n in inputs or i >= len(row) else cell_text(row[i])
                for n, i in positions.items()
            }
            item_id = values[self.id_column]
            if not item_id:
                continue  # blank/spacer row
            if item_id in seen:
                raise ItemSourceError(f"Duplicate item ID '{item_id}' in {self.name}.")
            seen.add(item_id)
            items.append(Item(id=item_id, values=values))
        return items
