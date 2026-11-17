"""Excel (.xlsx) implementation of :class:`~dlg.sources.base.ItemSource`."""

from __future__ import annotations

from collections.abc import Iterator
from pathlib import Path

from openpyxl import load_workbook

from dlg.models import Item

#: Default mapping of Item fields to Excel header names.
DEFAULT_COLUMNS: dict[str, str] = {
    "id": "ID",
    "name": "Name",
    "description": "Description",
    "category": "Category",
}


class ItemSourceError(Exception):
    """Raised when the item database is missing, malformed or inconsistent."""


class ExcelItemSource:
    """Reads items from the first sheet of an Excel workbook.

    The first row must contain headers. Columns not in the column map are kept
    in ``Item.extra``. Rows with an empty ID are skipped.

    Attributes:
        path: Location of the workbook.
        columns: Mapping of Item field name to Excel header name.
        name: Filename of the workbook.
    """

    def __init__(self, path: str | Path, columns: dict[str, str] | None = None) -> None:
        """Initialise the source.

        Args:
            path: Path to the .xlsx file.
            columns: Optional override of the field-to-header mapping.
        """
        self.path = Path(path)
        self.columns = columns or DEFAULT_COLUMNS
        self.name = self.path.name

    def load_items(self) -> list[Item]:
        """Load all items from the workbook.

        Returns:
            Items in sheet order.

        Raises:
            ItemSourceError: If the file is missing or unreadable, a required
                header is absent, or an ID appears more than once.
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
                return []
            index = self._header_index(header)
            return self._parse_rows(rows, index, header)
        finally:
            wb.close()

    def _header_index(self, header: tuple) -> dict[str, int]:
        """Map each Item field to its column position.

        Args:
            header: The first row of the sheet.

        Returns:
            Field name to zero-based column index.

        Raises:
            ItemSourceError: If a required header (ID, Name) is missing.
        """
        positions = {str(h).strip(): i for i, h in enumerate(header) if h is not None}
        index: dict[str, int] = {}
        for field, header_name in self.columns.items():
            if header_name in positions:
                index[field] = positions[header_name]
            elif field in ("id", "name"):
                raise ItemSourceError(
                    f"Required column '{header_name}' not found in {self.name}"
                )
        return index

    def _parse_rows(
        self, rows: Iterator[tuple], index: dict[str, int], header: tuple
    ) -> list[Item]:
        """Convert sheet rows to Items.

        Args:
            rows: Iterator of data rows (header already consumed).
            index: Field name to column index.
            header: The header row, used to name extra columns.

        Returns:
            Parsed items.

        Raises:
            ItemSourceError: On duplicate IDs.
        """
        mapped = set(index.values())
        items: list[Item] = []
        seen: set[str] = set()

        def cell(row: tuple, i: int) -> str:
            """Return a cell as a stripped string ('' if empty or out of range)."""
            if i >= len(row) or row[i] is None:
                return ""
            return str(row[i]).strip()

        for row in rows:
            item_id = cell(row, index["id"])
            if not item_id:
                continue  # blank/spacer row
            if item_id in seen:
                raise ItemSourceError(f"Duplicate item ID '{item_id}' in {self.name}")
            seen.add(item_id)
            extra = {
                str(header[i]).strip(): cell(row, i)
                for i in range(len(header))
                if i not in mapped and header[i] is not None and cell(row, i)
            }
            items.append(
                Item(
                    id=item_id,
                    name=cell(row, index["name"]),
                    description=cell(row, index["description"]) if "description" in index else "",
                    category=cell(row, index["category"]) if "category" in index else "",
                    extra=extra,
                )
            )
        return items
