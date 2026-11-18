# Changelog

Format: one entry per release, newest first. Versioning: [SemVer](https://semver.org/).

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
