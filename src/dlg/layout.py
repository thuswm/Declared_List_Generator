"""Word table layout drawn in an Excel sheet (multi-level header with merged cells).

Layout sheet conventions (default: a sheet called "Word layout" in the item
database workbook):

* The **last row in use is the mapping row**. Each of its cells names a column
  of the data sheet; the data rows go beneath the header in that order. The
  mapping row itself is not printed.
* The table spans the columns from the first to the last filled cell of the
  mapping row. A ``-`` there is a blank Word column on purpose (it may also
  sit at the table edge). An empty cell in between, or a name the database
  lacks, also gives a blank column, but with a warning.
* The rows above the mapping row (from the first row with text) are the Word
  header. They hold free text, and merged cells (across and down) become merged
  cells in Word. With no rows above, the mapping names form a one-row header.
* A mapping name ending in ``*`` is a user-input column. It does not have to
  exist in the database: the tool then adds it as an input column.
* Names match database columns exactly, else ignoring case and repeated,
  leading or trailing spaces (with a warning).
* Blank columns are Word-only: exported empty and not shown in the browser.
* The Excel column widths set the relative Word column widths.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from dlg.models import Column, Item, ItemTable
from dlg.sources.excel import INPUT_MARKER, cell_text

#: Default name of the layout sheet inside the item database workbook.
DEFAULT_LAYOUT_SHEET = "Word layout"

#: Mapping-cell text for a Word column that is blank on purpose.
BLANK_MARKER = "-"

#: Excel's default column width (characters) for columns without a set width.
EXCEL_DEFAULT_WIDTH = 8.43


class LayoutError(Exception):
    """Raised when the layout sheet cannot be used."""


@dataclass
class HeaderCell:
    """One (possibly merged) cell of the Word header.

    Attributes:
        row: Zero-based header row.
        col: Zero-based table column.
        text: Cell text (may contain line breaks).
        row_span: Number of header rows the cell covers.
        col_span: Number of columns the cell covers.
    """

    row: int
    col: int
    text: str
    row_span: int = 1
    col_span: int = 1

    @property
    def merged(self) -> bool:
        """True if the cell covers more than one grid cell."""
        return self.row_span > 1 or self.col_span > 1


@dataclass
class LayoutColumn:
    """One column of the Word table, as named in the mapping row.

    Attributes:
        name: Mapping text without the ``*`` ('' for an empty cell or ``-``).
        user_input: True if the name ended in ``*``.
        blank: True if the column is exported empty (``-``, an empty cell, or
            a name the database lacks).
        ref: Excel cell of the mapping name, e.g. ``C5`` (for messages).
    """

    name: str
    user_input: bool = False
    blank: bool = False
    ref: str = ""

    @property
    def label(self) -> str:
        """Display text: the name, with blank columns marked."""
        if not self.blank:
            return self.name
        return f"{self.name} (blank)" if self.name else "(blank)"


@dataclass
class TableLayout:
    """A Word table layout read from a layout sheet.

    Attributes:
        columns: Word columns, in table order (``user_input`` from a ``*``).
        header_rows: Number of header rows (at least 1).
        cells: Header cells covering the ``header_rows`` x ``len(columns)`` grid.
        widths: Relative column widths (Excel character units), one per column.
        source: Human-readable origin, e.g. ``items.xlsx › Word layout``.
        warnings: Problems that do not stop the export (blank columns that
            were probably not meant to be blank).
    """

    columns: list[LayoutColumn]
    header_rows: int
    cells: list[HeaderCell]
    widths: list[float]
    source: str = ""
    warnings: list[str] = field(default_factory=list)

    @property
    def column_names(self) -> list[str]:
        """Mapping names of the Word columns, in table order ('' for unnamed blanks)."""
        return [c.name for c in self.columns]

    @property
    def column_keys(self) -> list[str | None]:
        """Item value keys per Word column; None for a blank column."""
        return [None if c.blank else c.name for c in self.columns]

    @property
    def column_labels(self) -> list[str]:
        """Display texts of the Word columns, blank columns marked."""
        return [c.label for c in self.columns]


def load_layout(path: str | Path, sheet: str, required: bool) -> TableLayout | None:
    """Read the table layout from a workbook.

    Args:
        path: Workbook holding the layout sheet.
        sheet: Name of the layout sheet (matched ignoring case).
        required: True if the sheet must exist (a separate layout file was
            set). False for the default sheet inside the item database, where
            a missing sheet means "use the automatic table".

    Returns:
        The layout, or None if the sheet is absent and not required.

    Raises:
        LayoutError: If the file cannot be read, a required sheet is missing,
            or the sheet breaks one of the layout rules.
    """
    path = Path(path)
    if not path.is_file():
        raise LayoutError(f"Table layout file not found: {path}")
    try:
        # Not read-only: merged cells and column widths are only available in full mode.
        wb = load_workbook(path, data_only=True)
    except Exception as exc:  # openpyxl raises various types on bad files
        raise LayoutError(f"Cannot read table layout file {path}: {exc}") from exc
    try:
        ws = find_sheet(wb.sheetnames, sheet)
        if ws is None:
            if required:
                raise LayoutError(
                    f"Sheet '{sheet}' not found in {path.name}. Sheets: {', '.join(wb.sheetnames)}."
                )
            return None
        return parse_layout(wb[ws], f"{path.name} › {ws}")
    finally:
        wb.close()


def find_sheet(names: list[str], wanted: str) -> str | None:
    """Find a sheet by name: exact match first, then ignoring case.

    Args:
        names: Sheet names of the workbook.
        wanted: Name to look for.

    Returns:
        The actual sheet name, or None.
    """
    if wanted in names:
        return wanted
    folded = wanted.strip().casefold()
    return next((n for n in names if n.strip().casefold() == folded), None)


def parse_layout(ws: Worksheet, source: str) -> TableLayout:
    """Turn a layout worksheet into a :class:`TableLayout`.

    Args:
        ws: The layout sheet (loaded in full, not read-only, mode).
        source: Origin for messages, e.g. ``items.xlsx › Word layout``.

    Returns:
        The layout; an empty mapping cell inside the table becomes a blank
        column with a warning.

    Raises:
        LayoutError: If the sheet is empty, the mapping row has duplicates or
            a bare ``*``, or merged cells cross the mapping row or the table edge.
    """
    text = {
        (cell.row, cell.column): cell_text(cell.value)
        for row in ws.iter_rows()
        for cell in row
        if cell_text(cell.value)
    }
    if not text:
        raise LayoutError(f"Layout sheet {source} is empty: the last used row must name the columns.")
    map_row = max(r for r, _ in text)
    map_cols = sorted(c for r, c in text if r == map_row)
    first_col, last_col = map_cols[0], map_cols[-1]

    columns: list[LayoutColumn] = []
    warnings: list[str] = []
    for c in range(first_col, last_col + 1):
        ref = f"{get_column_letter(c)}{map_row}"
        name = text.get((map_row, c), "")
        if name == BLANK_MARKER:
            columns.append(LayoutColumn(name="", blank=True, ref=ref))
            continue
        if not name:
            # A gap is most likely a forgotten name, so it is exported blank but reported.
            warnings.append(
                f"Cell {ref} in the mapping row is empty: that Word column stays blank. "
                f"Type '{BLANK_MARKER}' there if it is meant to be blank."
            )
            columns.append(LayoutColumn(name="", blank=True, ref=ref))
            continue
        user_input = name.endswith(INPUT_MARKER)
        if user_input:
            name = name.rstrip(INPUT_MARKER).rstrip()
        if not name:
            raise LayoutError(f"Layout sheet {source}: cell {ref} has no column name before the '*'.")
        if name == BLANK_MARKER:
            raise LayoutError(
                f"Layout sheet {source}: cell {ref}: '{BLANK_MARKER}' marks a blank column "
                "and cannot be an input column."
            )
        if name in [col.name for col in columns]:
            raise LayoutError(f"Layout sheet {source}: column '{name}' is mapped twice (cell {ref}).")
        columns.append(LayoutColumn(name=name, user_input=user_input, ref=ref))

    in_table = [r for r, c in text if r < map_row and first_col <= c <= last_col]
    first_row = min(in_table, default=map_row)
    cells = _header_cells(ws, source, text, first_row, map_row, first_col, last_col)
    if not cells:
        # No header rows above the mapping row: the column names are the header.
        cells = [HeaderCell(0, i, col.name) for i, col in enumerate(columns)]
    return TableLayout(
        columns=columns,
        header_rows=max(map_row - first_row, 1),
        cells=cells,
        widths=_column_widths(ws, first_col, last_col),
        source=source,
        warnings=warnings,
    )


def _header_cells(
    ws: Worksheet,
    source: str,
    text: dict[tuple[int, int], str],
    first_row: int,
    map_row: int,
    first_col: int,
    last_col: int,
) -> list[HeaderCell]:
    """Build the header grid from the rows above the mapping row.

    Args:
        ws: The layout sheet.
        source: Origin for messages.
        text: Non-empty cell texts by (row, column), 1-based.
        first_row: First header row (1-based).
        map_row: The mapping row (1-based); header rows end above it.
        first_col: First table column (1-based).
        last_col: Last table column (1-based).

    Returns:
        Header cells with zero-based positions; empty when there are no header rows.

    Raises:
        LayoutError: If a merged range includes the mapping row or sticks out
            of the table's columns.
    """
    cells: list[HeaderCell] = []
    covered: set[tuple[int, int]] = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_row > map_row:
            continue  # empty merged cells below the table
        # Checked before the column test: merging clears every cell but the top-left one,
        # so a header merged down into the mapping row erases that column's name and the
        # table would silently start or end one column early.
        if rng.max_row >= map_row:
            raise LayoutError(
                f"Layout sheet {source}: merged cells {rng.coord} include the mapping row "
                f"(row {map_row}); each mapping cell must name exactly one column and not be merged."
            )
        if rng.max_col < first_col or rng.min_col > last_col or rng.max_row < first_row:
            continue  # unrelated merge elsewhere on the sheet (e.g. a note)
        if rng.min_col < first_col or rng.max_col > last_col or rng.min_row < first_row:
            raise LayoutError(
                f"Layout sheet {source}: merged cells {rng.coord} reach outside the table "
                f"(columns {get_column_letter(first_col)}–{get_column_letter(last_col)})."
            )
        cells.append(HeaderCell(
            row=rng.min_row - first_row,
            col=rng.min_col - first_col,
            text=text.get((rng.min_row, rng.min_col), ""),
            row_span=rng.max_row - rng.min_row + 1,
            col_span=rng.max_col - rng.min_col + 1,
        ))
        covered.update((r, c) for r in range(rng.min_row, rng.max_row + 1)
                       for c in range(rng.min_col, rng.max_col + 1))
    for r in range(first_row, map_row):
        for c in range(first_col, last_col + 1):
            if (r, c) not in covered:
                cells.append(HeaderCell(r - first_row, c - first_col, text.get((r, c), "")))
    cells.sort(key=lambda cell: (cell.row, cell.col))
    return cells


def _column_widths(ws: Worksheet, first_col: int, last_col: int) -> list[float]:
    """Read the Excel column widths of the table columns.

    openpyxl stores one dimension entry per *group* of equal columns (``min``
    to ``max``), and reports a made-up width for columns without one, so only
    entries with a custom width are used; the rest get Excel's default.

    Args:
        ws: The layout sheet.
        first_col: First table column (1-based).
        last_col: Last table column (1-based).

    Returns:
        One positive width per column.
    """
    default = ws.sheet_format.defaultColWidth or EXCEL_DEFAULT_WIDTH
    widths = {c: float(default) for c in range(first_col, last_col + 1)}
    for dim in list(ws.column_dimensions.values()):
        if not dim.customWidth or not dim.width or dim.min is None:
            continue
        for c in range(dim.min, (dim.max or dim.min) + 1):
            if c in widths:
                widths[c] = float(dim.width)
    return [widths[c] for c in range(first_col, last_col + 1)]


def apply_layout(table: ItemTable, layout: TableLayout) -> tuple[ItemTable, TableLayout]:
    """Merge the layout's columns into the item table shown in the browser.

    * A mapped column marked ``*`` becomes an input column even if the data
      sheet has it as a normal column (its sheet values are then ignored).
    * A mapped input column the database lacks is added (at the end) as an
      input column.
    * Any other mapped column the database lacks becomes a blank Word column,
      with a warning.
    * Names match exactly first, else ignoring case and repeated, leading or
      trailing spaces (with a warning naming the database column used).

    Args:
        table: Columns and items from the data sheet.
        layout: The layout.

    Returns:
        A new table and a new layout with unknown columns marked blank and
        their warnings added (the inputs are not modified).

    Raises:
        LayoutError: If the ID column is marked as user input, a name matches
            several database columns loosely, or two mapping names end up on
            the same column.
    """
    columns = [c.model_copy() for c in table.columns]
    by_name = {c.name: c for c in columns}
    word_columns: list[LayoutColumn] = []
    warnings = list(layout.warnings)
    used: dict[str, str] = {}  # database column -> mapping cell that took it
    for mapped in layout.columns:
        if mapped.blank:
            word_columns.append(mapped)
            continue
        existing = _match_column(mapped, by_name, layout.source)
        if existing is not None and existing.name != mapped.name:
            warnings.append(
                f"Column '{mapped.name}' (cell {mapped.ref}) matched the database column '{existing.name}' "
                "(case or spaces differ)."
            )
            mapped = replace(mapped, name=existing.name)
        if existing is not None:
            if existing.name in used:
                raise LayoutError(
                    f"Layout sheet {layout.source}: cells {used[existing.name]} and {mapped.ref} "
                    f"both map the column '{existing.name}'."
                )
            used[existing.name] = mapped.ref
        word_columns.append(mapped)
        if existing is None:
            if not mapped.user_input:
                available = ", ".join(c.name for c in table.columns) or "(none)"
                warnings.append(
                    f"Column '{mapped.name}' (cell {mapped.ref}) is not in the item database: that Word column "
                    f"stays blank. Available columns: {available}. (Add '*' to type it in the tool, "
                    f"or write '{BLANK_MARKER}' for a blank column.)"
                )
                word_columns[-1] = replace(mapped, blank=True)
                continue
            new = Column(name=mapped.name, user_input=True)
            columns.append(new)
            by_name[new.name] = new
            used[new.name] = mapped.ref
        elif mapped.user_input and not existing.user_input:
            if existing.name == table.id_column:
                raise LayoutError(
                    f"Layout sheet {layout.source}: the ID column '{existing.name}' cannot be an input column."
                )
            existing.user_input = True
    inputs = {c.name for c in columns if c.user_input}
    items = [
        Item(id=i.id, values={c.name: "" if c.name in inputs else i.values.get(c.name, "") for c in columns})
        for i in table.items
    ]
    return (
        ItemTable(id_column=table.id_column, columns=columns, items=items),
        replace(layout, columns=word_columns, warnings=warnings),
    )


def _normalised(name: str) -> str:
    """Fold a column name for lenient matching (case, repeated and outer spaces ignored).

    Args:
        name: Column name.

    Returns:
        The folded name.
    """
    return " ".join(name.split()).casefold()


def _match_column(mapped: LayoutColumn, by_name: dict[str, Column], source: str) -> Column | None:
    """Find the database column a mapping name refers to.

    Hand-typed layout names often differ from database headers only in case
    or spacing, so a loose match is accepted when it is unambiguous.

    Args:
        mapped: The mapping-row column.
        by_name: Known columns by exact name.
        source: Layout origin, for messages.

    Returns:
        The exact match, else the single loose match, else None.

    Raises:
        LayoutError: If several columns match loosely.
    """
    if mapped.name in by_name:
        return by_name[mapped.name]
    wanted = _normalised(mapped.name)
    loose = [c for name, c in by_name.items() if _normalised(name) == wanted]
    if len(loose) > 1:
        names = ", ".join(f"'{c.name}'" for c in loose)
        raise LayoutError(
            f"Layout sheet {source}: column '{mapped.name}' (cell {mapped.ref}) matches several "
            f"database columns ({names}); write the name exactly as in the database."
        )
    return loose[0] if loose else None
