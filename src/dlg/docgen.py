"""Word document generation for the declared list.

Opens a Word template, finds the placeholder paragraph and replaces it with a
table of the selected items. The table is inserted *in place*, so it inherits
the page orientation of whichever template section holds the placeholder (the
template is expected to put it in a landscape section).
"""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentType
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.section import Section
from docx.shared import Emu
from docx.table import Table, _Cell, _Row
from docx.text.paragraph import Paragraph

from dlg.models import Item

#: Text of the template paragraph that is replaced by the table.
PLACEHOLDER = "{{DECLARED_LIST}}"

# XML 1.0 forbids most control characters; Excel data can occasionally contain them.
_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


class DocgenError(Exception):
    """Raised when the document cannot be generated."""


@dataclass(frozen=True)
class Column:
    """One column of the declared-list table.

    Attributes:
        header: Header text.
        value: Function returning the cell text for a (1-based row number, item).
        weight: Relative width; widths are scaled to fill the section's text area.
    """

    header: str
    value: Callable[[int, Item], str]
    weight: float


#: The columns of the declared-list table. Edit here to change the output layout.
TABLE_COLUMNS: tuple[Column, ...] = (
    Column("No.", lambda n, i: str(n), 0.6),
    Column("ID", lambda n, i: i.id, 1.2),
    Column("Name", lambda n, i: i.name, 2.5),
    Column("Description", lambda n, i: i.description, 4.5),
    Column("Category", lambda n, i: i.category, 1.8),
)


def generate_declared_list(
    template_path: str | Path,
    items: Sequence[Item],
    output_path: str | Path,
    columns: Sequence[Column] = TABLE_COLUMNS,
) -> Path:
    """Create a declared list document and save it to disk.

    Args:
        template_path: Word template containing the placeholder paragraph.
        items: Items to list, in the order they should appear.
        output_path: Where to write the finished .docx.
        columns: Table column definitions.

    Returns:
        The output path.

    Raises:
        DocgenError: If the template is missing/invalid, has no placeholder,
            or no items were given.
    """
    doc = build_document(template_path, items, columns)
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out)
    return out


def render_to_bytes(
    template_path: str | Path,
    items: Sequence[Item],
    columns: Sequence[Column] = TABLE_COLUMNS,
) -> bytes:
    """Create a declared list document in memory (for HTTP downloads).

    Args:
        template_path: Word template containing the placeholder paragraph.
        items: Items to list, in order.
        columns: Table column definitions.

    Returns:
        The .docx file contents.

    Raises:
        DocgenError: See :func:`generate_declared_list`.
    """
    buf = BytesIO()
    build_document(template_path, items, columns).save(buf)
    return buf.getvalue()


def build_document(
    template_path: str | Path,
    items: Sequence[Item],
    columns: Sequence[Column] = TABLE_COLUMNS,
) -> DocumentType:
    """Load the template and replace the placeholder with the items table.

    Args:
        template_path: Word template containing the placeholder paragraph.
        items: Items to list, in order.
        columns: Table column definitions.

    Returns:
        The populated document (not yet saved).

    Raises:
        DocgenError: If the template is missing/invalid, has no placeholder,
            or ``items`` is empty.
    """
    if not items:
        raise DocgenError("No items selected; nothing to put in the declared list.")
    path = Path(template_path)
    if not path.is_file():
        raise DocgenError(f"Template not found: {path}")
    try:
        doc = Document(str(path))
    except Exception as exc:  # python-docx raises several types for bad packages
        raise DocgenError(f"Cannot read template {path.name}: {exc}") from exc

    placeholder = _find_placeholder(doc)
    section = doc.sections[_section_index(doc, placeholder._p)]
    table = _build_table(doc, items, columns, _text_width(section))

    # Insert *before* the placeholder: if it carries the section break, inserting
    # after it would push the table into the next section (wrong orientation).
    placeholder._p.addprevious(table._tbl)
    _remove_placeholder(placeholder)
    return doc


def _find_placeholder(doc: DocumentType) -> Paragraph:
    """Find the body paragraph whose text is exactly the placeholder.

    Args:
        doc: The template document.

    Returns:
        The placeholder paragraph.

    Raises:
        DocgenError: If no such paragraph exists.
    """
    for paragraph in doc.paragraphs:
        # Compare the joined text: Word often splits typed text across runs.
        if paragraph.text.strip() == PLACEHOLDER:
            return paragraph
    raise DocgenError(f"Template has no paragraph containing only {PLACEHOLDER}")


def _remove_placeholder(paragraph: Paragraph) -> None:
    """Remove the placeholder without losing a section break.

    If the placeholder is the last paragraph of a section, the section's page
    setup lives in this paragraph's properties; deleting it would merge the
    section into the next one. In that case only its text is cleared.

    Args:
        paragraph: The placeholder paragraph.
    """
    p = paragraph._p
    if p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
        for run in paragraph.runs:
            run._r.getparent().remove(run._r)
    else:
        p.getparent().remove(p)


def _section_index(doc: DocumentType, element: BaseOxmlElement) -> int:
    """Return the index in ``doc.sections`` of the section holding an element.

    A section's properties sit in the last paragraph of that section, so the
    index equals the number of section breaks that precede the element.

    Args:
        doc: The document.
        element: A body-level XML element (paragraph or table).

    Returns:
        Zero-based section index.
    """
    index = 0
    for child in doc.element.body:
        if child is element:
            return index
        if child.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
            index += 1
    return index  # pragma: no cover - element not in body


def _text_width(section: Section) -> int:
    """Return the usable text width of a section in EMU.

    Args:
        section: The section.

    Returns:
        Page width minus left/right margins and gutter.
    """
    return int(section.page_width - section.left_margin - section.right_margin)


def _clean(text: str) -> str:
    """Strip characters that cannot be stored in a Word document.

    Args:
        text: Raw cell text.

    Returns:
        Text safe to write into XML.
    """
    return _CONTROL_CHARS.sub("", text)


def _build_table(
    doc: DocumentType,
    items: Sequence[Item],
    columns: Sequence[Column],
    width: int,
) -> Table:
    """Create the items table (appended at the end; caller moves it).

    Args:
        doc: Document that will own the table.
        items: Items to list.
        columns: Column definitions.
        width: Total table width in EMU.

    Returns:
        The finished table with a repeating, bold, shaded header row.
    """
    table = doc.add_table(rows=1, cols=len(columns))
    _apply_borders(table)
    table.autofit = False

    total = sum(c.weight for c in columns)
    widths = [Emu(int(width * c.weight / total)) for c in columns]

    for cell, col in zip(table.rows[0].cells, columns):
        cell.text = col.header
        for run in cell.paragraphs[0].runs:
            run.bold = True
        _shade(cell, "D9D9D9")
    _mark_header_row(table.rows[0])

    for n, item in enumerate(items, start=1):
        row = table.add_row()
        _no_split(row)
        for cell, col in zip(row.cells, columns):
            cell.text = _clean(col.value(n, item))

    # Word honours per-cell widths but LibreOffice and others follow the column
    # grid, so set both, plus a fixed layout and total width, to render alike.
    for col, w in zip(table.columns, widths):
        col.width = w
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = w
    _fix_layout(table, sum(widths))
    return table


def _fix_layout(table: Table, total: int) -> None:
    """Set an explicit total table width (the fixed layout comes from ``autofit``).

    ``w:tblW`` already exists in tables made by python-docx; it is edited in
    place because OOXML requires ``w:tblPr`` children in a fixed order.

    Args:
        table: The table to adjust.
        total: Total table width in EMU.
    """
    tbl_w = table._tbl.tblPr.find(qn("w:tblW"))
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(int(Emu(total).twips)))


def _apply_borders(table: Table) -> None:
    """Give the table single-line borders regardless of template styles.

    Done in XML so it works even if the template lacks the "Table Grid" style.

    Args:
        table: The table to style.
    """
    tbl_pr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    # Schema order: ... tblBorders, shd, tblLayout, tblCellMar, tblLook.
    tbl_pr.insert_element_before(borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")


def _shade(cell: _Cell, fill: str) -> None:
    """Set a solid background colour on a cell.

    Args:
        cell: Table cell.
        fill: Hex colour without ``#``.
    """
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def _mark_header_row(row: _Row) -> None:
    """Make a row repeat at the top of every page the table spans.

    Args:
        row: The header row.
    """
    row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))


def _no_split(row: _Row) -> None:
    """Stop a row from breaking across two pages.

    Args:
        row: A data row.
    """
    row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
