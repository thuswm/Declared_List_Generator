# Decisions & Gotchas

Architecture decision log, newest first. Full per-phase history of release 0.1.0:
[archive/release-0.1.0-plan.md](archive/release-0.1.0-plan.md).

## 2026-11-18 – User-input columns (0.2.0 step 7)
- Header ending in `*` (trailing whitespace before it allowed, `Lot *` → `Lot`) marks an input column: `Column.user_input = True`. The sheet's cells in that column are **ignored** (values come only from the user), so `Item.values` holds `""` there. A header of only `*` is skipped; `X` + `X*` is a duplicate-header error; the ID column may not be an input column. `#X*` is simply omitted.
- Merging happens in one place per side: `ItemTable.with_user_values()` (server, export) and `withUserValues()` (GUI: table, search, filter lists). Both only fill input columns, so a session/request can never override database data.
- Export request carries `user_values`; the Word table shows them at the column's sheet position like any other column.
- GUI: input cells are `readOnly` (not `disabled`) for unselected rows so a click still reaches the cell: clicking one selects the row and focuses the input. Text-input `change` events are ignored by the selection handler (only checkboxes toggle). Typed values stay in memory when a row is deselected (reappear if reselected) but are dropped on save (selected items only).
- Example database has a `Serial number*` column.

## 2026-11-18 – Configuration name in the GUI (0.2.0 step 6)
- "Configuration" text field (max 500) under the session name; saved with the session (trimmed by the API), filled on load; editing it marks the session as changed. Enter saves.
- Saved list entry: `Name — Configuration — n item(s) — date` (configuration omitted when empty); full text also as tooltip. Header shows `Session: … · Configuration: …` after load/save, hidden after deleting the current session. Label builders `sessionLabel`/`currentLabel` live in `logic.js` (Node-tested).
- Saving under a new name keeps the field's configuration name ("save as" copies it).
- Gotcha (test tooling): `pkill -f "<pattern>"` also matched the shell running it, killed the shell and left a test server alive on the port, so a later run talked to the old server. Stop background test processes by PID (`$!`) instead.

## 2026-11-18 – Session schema v2 (0.2.0 step 5)
- `schema_version` 2 adds `configuration_name` (str) and `user_values` (`{item_id: {column: value}}`). `sessions.migrate()` upgrades raw data one version per step (v1 → v2 just sets the version; model defaults fill the new fields). The file on disk is only rewritten on the next save.
- On save, `user_values` keeps only non-empty values of *selected* items (values of deselected items are dropped), in selection order. API strips `configuration_name`.
- **Name-collision fix:** a session is identified by the exact `name` inside the file, not by the filename. Lookup tries `<safe name>.json` first and checks the stored name; otherwise scans the folder. New sessions get the first free filename (`x.json`, `x (2).json`, ...). `Path.exists()` follows the file system, so case-only differences collide correctly on Windows (separate files there, same as Linux). Old v1 files with sanitised names are still found by their real name.
- A corrupt file at the preferred path is reported (422/400) rather than silently skipped, as before. Unrelated corrupt files are skipped during scans.
- Windows device names (`CON`, `NUL`, `COM1`, ...) get a `_` prefix in the filename.
- GUI: session-name "overwrite?" check is now exact (case-sensitive), matching the server. Until steps 6–7, the GUI keeps the loaded session's `configuration_name`/`user_values` and sends them back on save, so nothing is lost.

## 2026-11-18 – Excel-style column filters (0.2.0 step 4)
- State: `filter.columns = {column: Set(allowed values)}`; a column is absent when unfiltered. Blank cells are the value `""`, listed as "(Blanks)" (last in the list); typing "blank" finds them.
- Like Excel: a column's value list only shows values from rows passing the *other* column filters (not its own, not search/selected-only). With text typed in the panel, OK keeps only the listed *and* ticked values. All values ticked = filter removed. OK is disabled when nothing is ticked (would hide everything).
- Value list sorted with natural order (`A2` before `A10`), case-insensitive.
- Filters are view state: not saved in sessions, reset on reload.
- Panel: one shared `role="dialog"` element, `position: fixed`, clamped inside the window; Escape/Cancel/outside click cancel; focus returns to the column's button. Closed on window resize (position would go stale).
- Pure logic in `logic.js` (`filterValues`, `matchesQuery`, `chooseFilter`, `filterItems`), Node-tested. UI checked in headless Chromium driven over the DevTools protocol (scratch script, not part of the repo: Playwright's Python package isn't installed).

## 2026-11-18 – Columns from the Excel file (0.2.0 step 3)
- Model: `ItemTable(id_column, columns: list[Column], items)`, `Item(id, values: dict[column -> text])`. `values` holds every shown column (incl. the ID column) in sheet order, `""` for empty cells. `Column` is an object (not a bare string) so step 7 can add a `user_input` flag without changing the API shape.
- `ItemSource.load() -> ItemTable` replaces `load_items()`. Fixed `name/description/category/extra` fields, `TABLE_COLUMNS`, the docgen `Column` dataclass and the category dropdown are gone.
- Headers are stripped; `#`-prefixed and empty headers are skipped; duplicate shown headers are an error (values are keyed by header). Missing ID column → error naming the `id_column` setting and listing the available columns. Empty sheet → error.
- Cell text: dates without time → `YYYY-MM-DD`; whole floats → no `.0`.
- Word column widths: weight = max(average text length, longest single word incl. header), clamped 4..40, scaled to the text width. Keeps codes unbroken and stops one long description column from squeezing the rest.
- Config: `id_column` (default `ID`, must be non-empty). Settings missing from an existing config file are **appended** with defaults (not rewritten, so user edits/comments survive), so new settings become visible to the user.
- Sessions are unchanged (IDs only), so v1 files keep working; if `id_column` is changed to a different column, old sessions' IDs will no longer match (reported as unknown on load).
- Example database gets a `#Internal note` column to demonstrate the omit marker.

## 2026-11-18 – README for Linux and Windows (0.2.0 step 2)
- Windows commands use `.venv\Scripts\python` directly (works in PowerShell and cmd without activation); `py -3` launcher for creating the venv.
- `".[dev]"` in double quotes: cmd does not treat single quotes as quoting.
- Gotcha: PowerShell may block `Activate.ps1` (execution policy); README offers calling the venv Python directly or `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.
- Windows paths in `config.toml`: doubled backslashes in `"..."`, or `'...'` literal strings, or forward slashes (checked with `tomllib`).

## 2026-11-18 – 0.2.0 scope update after Windows test
- Windows test: the exported `.docx` opens correctly in real MS Word.
- **Columns come from the Excel file** (names, number, order); nothing is hard-coded. The ID column name is a config setting `id_column` (default `ID`) because the real database's ID header is not known yet; IDs must be unique (sessions store selections by ID). Missing ID column → clear error naming the setting.
- **Omit marker:** a header starting with `#` (e.g. `#Internal note`) is ignored everywhere: not shown, not searchable, not exported. Empty headers are ignored as before.
- **User-input marker** stays a trailing `*`; markers are stripped for display/export.
- **Filtering:** the single category dropdown is replaced by Excel-style filters on every column (distinct-value checklist + text box), combined with the global search.
- Word table: Excel columns in sheet order, widths derived from content length (no per-column config). The Word table contains exactly the non-omitted Excel columns – the automatic `No.` column is dropped.
- README must give setup/run instructions for both Linux and Windows.
- Order: column model before user-input columns, since `*` columns build on it.

## 2026-11-17 – 0.2.0 design decisions
- **Config file:** TOML (read with `tomllib`, small hand-written writer), one per user, created with defaults on first run. Linux/macOS `$XDG_CONFIG_HOME/dlg/config.toml` (default `~/.config/dlg/config.toml`), Windows `%APPDATA%\dlg\config.toml`, override `--config <path>`. Replaces `DLG_*` env vars. Settings: `items_file`, `template_file`, `sessions_dir`, `output_dir`, `host`, `port` (+ `id_column` from 2026-11-18).
- **Folder/file browsing:** server-side browser dialog in the GUI (no `tkinter`), localhost only.
- **Output:** written to `output_dir` only (no browser download) under a user-given name (default from the configuration name); GUI shows the saved path. `{{CONFIGURATION_NAME}}` placeholder in the template.
- **User-input columns:** header ends with `*`; values per selected item saved in the session and written as extra Word columns.
- **Session schema v2:** `configuration_name`, `user_values` (`{item_id: {column: value}}`); v1 files still load; filename-collision fix (`a/b` vs `a_b`) in the same change.
- **Exit button:** `POST /api/shutdown`, localhost only; GUI shows "server stopped".

## 2026-11-17 – Versioning and documentation layout
- Semantic versioning, annotated git tags (`v0.1.0` = first workable version). Version lives in `pyproject.toml` and matches the tag. Session `schema_version` is separate from the app version.
- Short-lived feature branches merged into `main`; tag `main` at each release.
- `PLAN.md` holds only the current release plus links; finished phases move to `docs/archive/` at release time.
- `my_notes.md` is the user's private scratch file (git-ignored); decisions from it are copied into these docs.
- The tool is developed on Linux but must also work on Windows; each user has a personal config file, created on first run or when missing.

## 2026-11-17 – Config file (0.2.0 step 1)
- `src/dlg/config.py`: `Config` dataclass, `load_config` (creates defaults if missing), `save_config` (atomic, hand-written flat TOML via JSON string escaping so Windows backslashes are safe), `default_config_path` (XDG/`~/.config` or `%APPDATA%`).
- Relative paths in the file resolve against the config file's folder; defaults are absolute paths in the project folder (so the file written on first run is explicit). Unknown keys are ignored; wrong types / bad TOML raise `ConfigError` (launcher prints it, exit code 2).
- `Settings` and the `DLG_*` env vars are gone. The module-level `app = create_app()` was removed because importing `dlg.app` would create a config file as a side effect; the launcher passes `create_app(config)` to uvicorn (so `uvicorn dlg.app:app` no longer works; use `python -m dlg`). Tests never touch the real user config.

## Carried over from 0.1.0 (still relevant)
- **Stack:** FastAPI + pydantic + uvicorn, vanilla HTML/JS, `openpyxl` (read), `python-docx` (write).
- **Item DB** sits behind the `ItemSource` interface (Excel in v1). IDs are always strings; duplicate IDs / missing ID or Name headers raise `ItemSourceError`; unmapped columns go to `Item.extra`.
- **Sessions:** JSON validated by pydantic, atomic writes, newer `schema_version` refused. (Name-collision gotcha fixed in 0.2.0 step 5.)
- **Word export:** placeholder paragraph must be exactly `{{DECLARED_LIST}}` in a landscape section; columns defined in `TABLE_COLUMNS`; rows follow database order. Gotchas: keep the section break when the placeholder paragraph holds it; set `tblGrid` widths (LibreOffice); keep `w:tblPr` child order (Word). Not yet verified in real Word.
- **API:** Excel re-read on every request; export rejects empty (400) / unknown IDs (422). Session routes use `{name:path}`.
- **GUI:** pure logic in `logic.js` (Node-testable); data written with `textContent` only (no `innerHTML`). Unknown IDs in a loaded session are dropped with a warning.
- **Security:** default bind is 127.0.0.1 because there is no authentication.
