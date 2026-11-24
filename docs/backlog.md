# Backlog

Source of truth for ideas not in the current release plan. The user keeps a private `my_notes.md`; its decisions are copied here.

## Scheduled
Nothing scheduled; next release to be planned in `PLAN.md`.

## Future ideas
- **Document review mode:** read a Word file, compare its item table with the database (Excel), highlight rows/cells that do not match.
- **Session library follow-ups** (after the user's try-out of 0.3.0): should plain Save write back to the folder the session was loaded from? Is the folder dialog the right way to pick the Save to… target? Open single `.json` session files directly (dropped from 0.3.0).
- **Merge several sessions** into one selection (user finds it interesting, wants to think it through first: conflicts in user values, configuration name, order).
- Swap Excel for another `ItemSource` (e.g. SQLite).
- Cache the Excel read if the database grows large.

## Known issues / checks
- **0.5.0 open checks (Windows, real Word):** `\blank` (several, any case) and an old `-` layout; ID column pick list (change the database path before saving, an unreadable file, a saved ID the database lacks); **Open** for database/template/layout/output folder with Excel, Word and Explorer (and the "Save settings first" message); rotated header text (−90° and +90°, merged down, with and without a hand-set Excel row height) in real Word and in *Check Word table*; width shares in *Check Word table*. Also confirm the 0.4.0 `-` problem is gone with the original layout file.
- **0.4.0 open checks (Windows):** a layout sheet with an empty mapping cell, a `-` column (also at the edge), an unknown name and a name in different case/spacing → warnings, ⚠ label, blank Word columns; *Check Word table…* dialog; Settings → *Empty input cells* (TBD in Word) and the export summary; session library dates as `YYYY-MM-DD HH:MM`; `start_dlg.bat` by double-click and via a desktop shortcut (also a path with spaces), and its error window when `.venv` is missing.
- **0.3.0 open checks:** try the session library, Save to…, a layout sheet (start from `data/layout_example.xlsx`) in real Word, the table font size, and Open document/folder with real Word and Explorer on Windows.
- **0.2.0 Windows test (done 2026-11-19 by the user):** works in general. Findings → 0.3.0 step 1: page needed F5 several times to load the new version; file-name field snapped back to the old name when the last character was deleted; "Saved" drop-down is not searchable (→ session browser); Word table needs a three-level header (→ step 5).
- Real Word template from the user: placeholder `{{DECLARED_LIST}}` in a landscape section.
- Many columns may not fit the landscape width: table font size and layout-sheet widths added in 0.3.0; check with the real database.
