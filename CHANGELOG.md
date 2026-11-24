# Changelog

Format: one entry per release, newest first. Versioning: [SemVer](https://semver.org/).

## [Unreleased]
- **`\blank` marks a blank Word column** in the layout sheet's mapping row (any case, as often as needed). `-` still works, as do dashes such as `–` that autocorrect puts in its place (likely cause of "`-` does not work" in 0.4.0).

## [0.4.0] – 2026-11-20
Changes from the Windows test of 0.3.0, plus four smaller additions. Not yet tried on Windows.
- **Blank Word columns instead of layout errors:** an empty cell in the layout sheet's mapping row, or a name the database lacks, now gives an empty Word column and a **warning** (status bar, ⚠ on the "Word table" label) instead of stopping. A **`-`** in the mapping row is a blank column on purpose (no warning, also at the table edge).
- **Lenient column names:** mapping names match database columns ignoring case and extra spaces, with a warning. Ambiguous matches and two names on one column are errors.
- **"TBD" for empty input cells:** new settings `mark_empty_inputs` / `empty_input_text` (Settings → *Empty input cells*) write a text such as `TBD` into empty cells of input (`*`) columns in Word.
- **Export summary:** after every export the status lists the empty input cells per column.
- **Check Word table…:** a dialog shows the Word header as it will be exported (merges, widths, column sources) plus layout warnings, before opening Word.
- **ISO dates** (`YYYY-MM-DD HH:MM`) in the session library, independent of the browser language.
- **`start_dlg.bat`:** start the tool on Windows by double-click or a desktop shortcut.
- Database columns that the layout does not map are documented as allowed. Session files unchanged (`schema_version` 2).

## [0.3.0] – 2026-11-19
Changes from the Windows test of 0.2.0. Steps 3–7 not yet tried on Windows; Open document/folder not yet tried with real Word/Explorer.
- **No more F5 after an update:** the GUI files are always revalidated; the app version is shown in the page header.
- **File-name field** no longer snaps back to the old name when emptied; the suggestion returns when the field is left empty, or with **Reset name**.
- **Word template** chosen in Settings with Browse.
- **Session library** dialog replaces the "Saved" drop-down: My sessions, a **Shared library** (new setting `global_sessions_dir`) or any folder; search, sorting, load and delete.
- **Save to…** saves the session into a chosen folder (e.g. the shared library) and asks before replacing a file there.
- **Multi-level Word table header** drawn in an Excel layout sheet (`Word layout` in the database workbook, or settings `table_layout_file` / `table_layout_sheet`): free header text, merged cells across and down, last row maps the database columns (`*` = input column, may exist only in the layout), Excel column widths become the Word widths. Example: `data/layout_example.xlsx`. Without a layout sheet the table is built as before.
- **Table font size** setting (`table_font_size`, 0 = template's size); with a size set the cells get compact spacing.
- **Open document / Open folder** buttons after an export (local only; only the exported file in the output folder).
- `output/` is git-ignored. Session files unchanged (`schema_version` 2).

## [0.2.0] – 2026-11-18
Changes from the first test run on Windows. Not yet tested on Windows or in real Word.
- **Per-user config file** (`config.toml`, created on first run; `%APPDATA%\dlg` on Windows, `~/.config/dlg` on Linux) replaces the `DLG_*` environment variables. New settings are appended to older files.
- **Columns come from the Excel file** (names, number, order) for the table and the Word export. ID column set by `id_column` (default `ID`). Header `#Name` = column ignored; header `Name*` = value typed by the user per item.
- **Excel-style column filters** on every column (value checklist with search), replacing the category dropdown.
- **User-input columns:** editable cells for selected items, saved in the session and exported.
- **Configuration name** per session: field in the GUI, shown in the saved list and page header, `{{CONFIGURATION_NAME}}` placeholder in the template.
- **Settings panel** with a Browse dialog for the item database and output folder, plus the ID column; saved to the config file.
- **Export** writes to the output folder under a chosen file name (suggested from the configuration name), asks before overwriting, refuses names that are invalid on Windows, reports files left open in Word.
- **Exit button** stops the server cleanly.
- **Session files v2** (`configuration_name`, `user_values`); v1 files still load. Names that map to the same file name (`a/b`, `a_b`, case-only on Windows) no longer overwrite each other.
- README: setup and run instructions for Windows and Linux.

## [0.1.0] – 2026-11-17
First workable version.
- Browser GUI: searchable/filterable item table, select-all-visible, save/load/delete sessions, export.
- Item database read from an Excel file (`ItemSource` interface).
- Sessions stored as JSON (`schema_version` 1).
- Word export of the selection as a table in a landscape section of a template.
- `python -m dlg` launcher; README.
