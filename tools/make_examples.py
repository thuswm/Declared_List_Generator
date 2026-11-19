"""Generate the example item database, Word table layout and Word template.

Run from the project root: ``.venv/bin/python tools/make_examples.py``
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.shared import Mm
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent

#: (ID, Name, Description, Category, Supplier)
SAMPLE_ITEMS: list[tuple[str, str, str, str, str]] = [
    ("DL-001", "Hydraulic Pump", "Main hydraulic pump assembly, 200 bar", "Hydraulics", "Acme Fluid"),
    ("DL-002", "Pressure Sensor", "0-250 bar analogue pressure transducer", "Sensors", "SensorTech"),
    ("DL-003", "Temperature Probe", "PT100 probe, 3-wire, stainless sheath", "Sensors", "SensorTech"),
    ("DL-004", "Control Unit", "PLC-based control unit, 24 V DC", "Electronics", "Ctrl Systems"),
    ("DL-005", "Cable Harness A", "Main power and signal harness", "Electrical", "WireWorks"),
    ("DL-006", "Cable Harness B", "Auxiliary harness for sensor array", "Electrical", "WireWorks"),
    ("DL-007", "Relief Valve", "Adjustable pressure relief valve, 50-300 bar", "Hydraulics", "Acme Fluid"),
    ("DL-008", "Display Panel", "7-inch industrial touch display", "Electronics", "Ctrl Systems"),
    ("DL-009", "Mounting Bracket", "Stainless steel mounting bracket, type 2", "Mechanical", "MetalCraft"),
    ("DL-010", "Gasket Set", "Complete gasket and O-ring service set", "Mechanical", "MetalCraft"),
    ("DL-011", "Filter Element", "10 micron return line filter element", "Hydraulics", "Acme Fluid"),
    ("DL-012", "Power Supply", "24 V DC, 10 A DIN-rail power supply", "Electrical", "WireWorks"),
]


def make_items_xlsx(path: Path) -> None:
    """Write the example item database.

    Args:
        path: Destination .xlsx path.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Items"
    # "Serial number*" shows the user-input marker (typed in the tool, cells left empty);
    # "#Internal note" shows the omit marker (the tool ignores that column).
    ws.append(["ID", "Name", "Description", "Category", "Supplier", "Serial number*", "#Internal note"])
    for row in SAMPLE_ITEMS:
        ws.append([*row, None, f"Checked by purchasing ({row[0]})"])
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def make_layout_xlsx(path: Path) -> None:
    """Write an example Word table layout for the example database.

    Kept in its own workbook (chosen with "Table layout" in Settings) so the
    example database still shows the automatic table out of the box. Copy the
    sheet into the database workbook to make it the default there.

    Args:
        path: Destination .xlsx path.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Word layout"
    ws.append(["Item", None, "Supplier", "Category", "Traceability", None])
    ws.append(["No.", "Designation", None, None, "Serial number", "Checked\nby"])
    # Mapping row (not printed): data-sheet columns; "Checked by*" exists only here (typed in the tool).
    ws.append(["ID", "Name", "Supplier", "Category", "Serial number*", "Checked by*"])
    # Across (Item, Traceability) and down (Supplier, Category span both header rows).
    for rng in ("A1:B1", "C1:C2", "D1:D2", "E1:F1"):
        ws.merge_cells(rng)
    for letter, width in zip("ABCDEF", (10, 30, 16, 14, 18, 12)):
        ws.column_dimensions[letter].width = width
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def make_template_docx(path: Path) -> None:
    """Write the example Word template.

    Layout: portrait intro page (with ``{{CONFIGURATION_NAME}}``), a landscape
    section holding the ``{{DECLARED_LIST}}`` placeholder, then portrait
    closing page(s).

    Args:
        path: Destination .docx path.
    """
    doc = Document()

    first = doc.sections[0]
    first.page_width, first.page_height = Mm(210), Mm(297)  # A4 portrait
    doc.add_heading("Declared List", level=0)
    # Filled with the configuration name entered in the tool.
    doc.add_paragraph("Configuration: {{CONFIGURATION_NAME}}")
    doc.add_paragraph(
        "This document contains the declared list for the project. "
        "The list is provided in table form on the following page(s)."
    )

    # New sections inherit the previous section's setup, so only swap what changes.
    landscape = doc.add_section(WD_SECTION.NEW_PAGE)
    landscape.orientation = WD_ORIENT.LANDSCAPE
    landscape.page_width, landscape.page_height = Mm(297), Mm(210)
    doc.add_heading("Declared Items", level=1)
    doc.add_paragraph("{{DECLARED_LIST}}")

    portrait = doc.add_section(WD_SECTION.NEW_PAGE)
    portrait.orientation = WD_ORIENT.PORTRAIT
    portrait.page_width, portrait.page_height = Mm(210), Mm(297)
    doc.add_heading("Notes", level=1)
    doc.add_paragraph("Any items not listed above are not part of this declaration.")

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


if __name__ == "__main__":
    make_items_xlsx(ROOT / "data" / "items.xlsx")
    print("Wrote data/items.xlsx")
    make_layout_xlsx(ROOT / "data" / "layout_example.xlsx")
    print("Wrote data/layout_example.xlsx")
    make_template_docx(ROOT / "templates" / "declared_list_template.docx")
    print("Wrote templates/declared_list_template.docx")
