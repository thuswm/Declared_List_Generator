"""Word document generation for the declared list.

Opens a Word template, finds the placeholder paragraph and replaces it with a
table of the selected items. The table is inserted *in place*, so it inherits
the page orientation of whichever template section holds the placeholder (the
template is expected to put it in a landscape section). Text placeholders such
as ``{{CONFIGURATION_NAME}}`` are filled in anywhere in the body, in tables and
in headers/footers.
"""

from __future__ import annotations

import os
import re
from collections.abc import Iterator, Sequence
from pathlib import Path

from docx import Document
from docx.document import Document as DocumentType
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.oxml.xmlchemy import BaseOxmlElement
from docx.section import Section
from docx.shared import Emu, Pt
from docx.table import Table, _Cell, _Row
from docx.text.paragraph import Paragraph

from dlg.layout import HeaderCell, TableLayout
from dlg.models import Item

#: Text of the template paragraph that is replaced by the table.
PLACEHOLDER = "{{DECLARED_LIST}}"

#: Replaced by the configuration name wherever it appears in the template.
CONFIGURATION_PLACEHOLDER = "{{CONFIGURATION_NAME}}"

# XML 1.0 forbids most control characters; Excel data can occasionally contain them.
_CONTROL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


class DocgenError(Exception):
    """Raised when the document cannot be generated."""


class OutputLockedError(DocgenError):
    """Raised when the output file cannot be replaced (e.g. it is open in Word)."""


#: Limits for a column's relative width (in characters), so one long text
#: column cannot squeeze the others to nothing and tiny columns stay readable.
MIN_WEIGHT = 4.0
MAX_WEIGHT = 40.0


def column_weights(columns: Sequence[str], items: Sequence[Item]) -> list[float]:
    """Estimate relative column widths from the content.

    A column gets the larger of its average text length and its longest single
    word (so IDs and part numbers are not broken mid-word), clamped to
    ``MIN_WEIGHT``..``MAX_WEIGHT``. Widths are later scaled to the page.

    Args:
        columns: Column names (the header texts).
        items: Items whose values fill the table.

    Returns:
        One weight per column.
    """
    weights = []
    for name in columns:
        texts = [i.values.get(name, "") for i in items]
        average = sum(len(t) for t in texts) / len(texts) if texts else 0.0
        words = [w for t in [name, *texts] for w in t.split()]
        longest_word = max((len(w) for w in words), default=0)
        weights.append(min(max(average, longest_word, MIN_WEIGHT), MAX_WEIGHT))
    return weights


def fill_empty_inputs(
    items: Sequence[Item], columns: Sequence[str], text: str
) -> tuple[list[Item], dict[str, int]]:
    """Write a text (e.g. ``TBD``) into empty cells of the given input columns.

    Done on the items rather than in the table builder so the layout table and
    the automatic table behave alike, and blank layout columns (which have no
    item key) are never touched.

    Args:
        items: Items with the user's values filled in.
        columns: Input columns that are in the Word table.
        text: Text for empty cells (cells holding only spaces count as empty).

    Returns:
        New items (the inputs are not modified) and the number of filled
        cells per column, for every column given (0 included), in order.
    """
    counts = {name: 0 for name in columns}
    filled: list[Item] = []
    for item in items:
        values = dict(item.values)
        for name in columns:
            if not values.get(name, "").strip():
                values[name] = text
                counts[name] += 1
        filled.append(Item(id=item.id, values=values))
    return filled, counts


def generate_declared_list(
    template_path: str | Path,
    columns: Sequence[str],
    items: Sequence[Item],
    output_path: str | Path,
    configuration_name: str = "",
    layout: TableLayout | None = None,
    font_size: float = 0,
) -> Path:
    """Create a declared list document and save it to disk.

    The document is written to a temporary file next to the target and then
    moved into place, so an existing file is never left half-written.

    Args:
        template_path: Word template containing the placeholder paragraph.
        columns: Names of the columns to put in the table, in order.
        items: Items to list, in the order they should appear.
        output_path: Where to write the finished .docx (folders are created).
        configuration_name: Text for ``{{CONFIGURATION_NAME}}``.
        layout: Multi-level header and column order from the layout sheet;
            None for the automatic one-row header over ``columns``.
        font_size: Font size (pt) for the whole table; 0 keeps the template's.

    Returns:
        The output path.

    Raises:
        OutputLockedError: If the target exists and cannot be replaced
            (typically because it is open in Word on Windows).
        DocgenError: If the template is missing/invalid, has no placeholder,
            no items were given, or the file cannot be written.
    """
    doc = build_document(template_path, columns, items, configuration_name, layout, font_size)
    out = Path(output_path)
    tmp = out.with_name(f"~{out.stem}.tmp.docx")
    try:
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(tmp))
        os.replace(tmp, out)
    except PermissionError as exc:
        tmp.unlink(missing_ok=True)
        raise OutputLockedError(
            f"Cannot write {out.name}: it may be open in Word. Close it and try again."
        ) from exc
    except OSError as exc:
        tmp.unlink(missing_ok=True)
        raise DocgenError(f"Cannot write {out}: {exc.strerror or exc}") from exc
    return out


def build_document(
    template_path: str | Path,
    columns: Sequence[str],
    items: Sequence[Item],
    configuration_name: str = "",
    layout: TableLayout | None = None,
    font_size: float = 0,
) -> DocumentType:
    """Load the template, fill in text placeholders and insert the items table.

    Args:
        template_path: Word template containing the placeholder paragraph.
        columns: Names of the columns to put in the table, in order.
        items: Items to list, in order.
        configuration_name: Text for ``{{CONFIGURATION_NAME}}``.
        layout: Header and column order from the layout sheet (``columns``
            is then ignored); None for the automatic table.
        font_size: Font size (pt) for the whole table; 0 keeps the template's.

    Returns:
        The populated document (not yet saved).

    Raises:
        DocgenError: If the template is missing/invalid, has no placeholder,
            or ``items`` or ``columns`` is empty.
    """
    if not items:
        raise DocgenError("No items selected; nothing to put in the declared list.")
    keys: Sequence[str | None] = columns if layout is None else layout.column_keys
    if not keys:
        raise DocgenError("The item database has no columns to export.")
    doc = _open_template(template_path)

    # Before the table exists, so item text that happens to contain a
    # placeholder is never substituted.
    for paragraph in _all_paragraphs(doc):
        _replace_in_paragraph(paragraph, CONFIGURATION_PLACEHOLDER, _clean(configuration_name))

    placeholder = _find_placeholder(doc)
    section = doc.sections[_section_index(doc, placeholder._p)]
    table = _build_table(doc, items, keys, _text_width(section), layout)
    if font_size:
        _set_font_size(table, font_size)

    # Insert *before* the placeholder: if it carries the section break, inserting
    # after it would push the table into the next section (wrong orientation).
    placeholder._p.addprevious(table._tbl)
    _remove_placeholder(placeholder)
    return doc


def _paragraphs_in(container: object) -> Iterator[Paragraph]:
    """Yield the paragraphs of a body, header, footer or cell, including nested tables.

    Args:
        container: Any python-docx object with ``paragraphs`` and ``tables``.

    Yields:
        Each paragraph.
    """
    yield from container.paragraphs
    for table in container.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from _paragraphs_in(cell)


def _all_paragraphs(doc: DocumentType) -> Iterator[Paragraph]:
    """Yield every paragraph of the body and of all headers and footers.

    Headers/footers linked to the previous section are skipped: they have no
    content of their own, and touching them would create an empty one.

    Args:
        doc: The document.

    Yields:
        Each paragraph.
    """
    yield from _paragraphs_in(doc)
    for section in doc.sections:
        for part in (
            section.header, section.first_page_header, section.even_page_header,
            section.footer, section.first_page_footer, section.even_page_footer,
        ):
            if not part.is_linked_to_previous:
                yield from _paragraphs_in(part)


def _replace_in_paragraph(paragraph: Paragraph, old: str, new: str) -> None:
    """Replace text in a paragraph, keeping run formatting where possible.

    Word often splits typed text across several runs. A placeholder inside one
    run is replaced there (keeping its formatting); one spread over several
    runs is handled by moving the paragraph's text into its first run.

    Args:
        paragraph: The paragraph.
        old: Text to find.
        new: Replacement text.
    """
    runs = paragraph.runs
    if old not in "".join(r.text for r in runs):
        return
    for run in runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
    joined = "".join(r.text for r in runs)
    if old in joined:
        runs[0].text = joined.replace(old, new)
        for run in runs[1:]:
            run.text = ""


def check_template(template_path: str | Path) -> None:
    """Check that a file can be used as the Word template.

    Used when the template is chosen in Settings, so a wrong file is reported
    then and not only at the next export.

    Args:
        template_path: Candidate template file.

    Raises:
        DocgenError: If the file is missing, not a readable Word document, or
            has no ``{{DECLARED_LIST}}`` placeholder paragraph.
    """
    _find_placeholder(_open_template(template_path))


def _open_template(template_path: str | Path) -> DocumentType:
    """Open the Word template.

    Args:
        template_path: Template file.

    Returns:
        The loaded document.

    Raises:
        DocgenError: If the file is missing or not a readable Word document.
    """
    path = Path(template_path)
    if not path.is_file():
        raise DocgenError(f"Template not found: {path}")
    try:
        return Document(str(path))
    except Exception as exc:  # python-docx raises several types for bad packages
        raise DocgenError(f"Cannot read template {path.name}: {exc}") from exc


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
    columns: Sequence[str | None],
    width: int,
    layout: TableLayout | None = None,
) -> Table:
    """Create the items table (appended at the end; caller moves it).

    Args:
        doc: Document that will own the table.
        items: Items to list.
        columns: Column names (keys into ``Item.values``; also the header
            texts when there is no layout). None marks a blank layout column.
        width: Total table width in EMU.
        layout: Multi-level header and relative widths; None for a one-row
            header with widths estimated from the content.

    Returns:
        The finished table with repeating, bold, shaded header row(s).
    """
    header_rows = layout.header_rows if layout else 1
    table = doc.add_table(rows=header_rows, cols=len(columns))
    _apply_borders(table)
    table.autofit = False

    weights = layout.widths if layout else column_weights(columns, items)
    total = sum(weights)
    widths = [Emu(int(width * w / total)) for w in weights]

    for item in items:
        row = table.add_row()
        _no_split(row)
        for cell, name in zip(row.cells, columns):
            cell.text = "" if name is None else _clean(item.values.get(name, ""))

    # Word honours per-cell widths but LibreOffice and others follow the column
    # grid, so set both, plus a fixed layout and total width, to render alike.
    # Set before merging: python-docx adds up the widths of merged cells.
    for col, w in zip(table.columns, widths):
        col.width = w
    for row in table.rows:
        for cell, w in zip(row.cells, widths):
            cell.width = w

    cells = layout.cells if layout else [HeaderCell(0, i, name) for i, name in enumerate(columns)]
    for spec in cells:
        cell = table.cell(spec.row, spec.col)
        if spec.merged:
            cell = cell.merge(table.cell(spec.row + spec.row_span - 1, spec.col + spec.col_span - 1))
        _fill_header_cell(cell, spec.text, centred=spec.merged)
    for row in table.rows[:header_rows]:
        _mark_header_row(row)
    _fix_layout(table, sum(widths))
    return table


def _set_font_size(table: Table, size: float) -> None:
    """Set one font size for all text in the table, with compact cell spacing.

    The size is set on every run *and* on every paragraph mark. Word sizes an
    empty paragraph (an empty cell) by its mark, so without it empty cells keep
    the template's larger font and make their rows taller than the rest.

    Paragraphs in the cells also get no space before/after and single line
    spacing: templates often add space after every paragraph (e.g. 10 pt at
    1.15 lines), which otherwise dominates the row height once the text is
    small (user decision: only when a size is set, so the template look is
    kept by default).

    Args:
        table: The finished items table.
        size: Font size in points (half-point steps).
    """
    half_points = str(int(round(size * 2)))
    for tc in table._tbl.iter(qn("w:tc")):  # each (merged) cell once
        for p in tc.iter(qn("w:p")):
            paragraph = Paragraph(p, table)
            for run in paragraph.runs:
                run.font.size = Pt(size)
            fmt = paragraph.paragraph_format  # python-docx keeps pPr children in schema order
            fmt.space_before = Pt(0)
            fmt.space_after = Pt(0)
            fmt.line_spacing = 1.0
            p_pr = p.get_or_add_pPr()
            r_pr = p_pr.find(qn("w:rPr"))
            if r_pr is None:
                # rPr sits near the end of pPr (only sectPr/pPrChange follow, never in table cells).
                r_pr = OxmlElement("w:rPr")
                p_pr.append(r_pr)
            for tag in ("w:sz", "w:szCs"):
                el = r_pr.find(qn(tag))
                if el is None:
                    el = OxmlElement(tag)
                    r_pr.append(el)
                el.set(qn("w:val"), half_points)


def _fill_header_cell(cell: _Cell, text: str, centred: bool) -> None:
    """Write a header cell: bold text on grey, vertically centred.

    Args:
        cell: The (possibly merged) header cell.
        text: Its text; line breaks from Excel (Alt+Enter) are kept.
        centred: Centre the text horizontally (used for merged cells, which
            usually title the columns beneath them).
    """
    cell.text = _clean(text)  # python-docx turns "\n" into a line break
    for paragraph in cell.paragraphs:
        if centred:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.bold = True
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _shade(cell, "D9D9D9")


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
