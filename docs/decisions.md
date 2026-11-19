# Decisions & Gotchas

Architecture decision log, newest first. Full per-phase history of release 0.1.0:
[archive/release-0.1.0-plan.md](archive/release-0.1.0-plan.md).

## 2026-11-19 – Table font size (0.3.0 step 6)
- Setting `table_font_size` (pt, a TOML number; `0` = keep the template's size), valid range 4–72 in half-point steps (`config.font_size_problem`, shared by the config loader and `PUT /api/settings`). GUI: a number field in Settings; empty = 0. The GUI refuses `validity.badInput` itself, because `Number('')`/NaN would otherwise turn a typo into "template font" (and `JSON.stringify(NaN)` is `null`, which the API reads as "unchanged").
- `docgen._set_font_size` sets the size on every run **and every paragraph mark** (`w:pPr/w:rPr/w:sz` + `szCs`). Word sizes an empty cell by its paragraph mark, so empty cells would otherwise keep the template size and make rows taller. It walks `w:tc` elements so merged cells are handled once. `rPr` is appended to `pPr` after `jc` (schema order; only `sectPr`/`pPrChange` may follow, never in cells).
- **Compact cells (user decision, option 2 of 3):** the example template's document defaults are 10 pt space after each paragraph and 1.15 line spacing. These apply inside the table, so at 7 pt the rows were barely lower (seen in the LibreOffice rendering). When a font size is set, cell paragraphs therefore also get 0 space before/after and single line spacing. With size 0 the template's spacing is kept untouched. The rejected options were "always compact" and "leave it to the template".

## 2026-11-19 – Word table layout sheet (0.3.0 step 5)
- **User decisions:** the mapping row is **not printed** (only the rows above it form the header, so a title can span down through all levels; with no rows above, the mapping names are a one-row header). The header style stays fixed: bold, grey, vertically centred, and merged cells also centred horizontally. It is not copied from Excel.
- `dlg/layout.py`: `load_layout(path, sheet, required)` → `TableLayout(columns, header_rows, cells: HeaderCell(row, col, text, row_span, col_span), widths, source)`. The workbook is opened in **full mode** (read-only mode has no merged cells or column widths). The sheet name is matched exactly first, then ignoring case.
- **Rules:** the table spans the first to the last filled cell of the mapping row, and gaps are an error. Header rows run from the first row with text in those columns. Unrelated merges outside the table are ignored. A merge that partly sticks out of the table is an error, and so is **any merge touching the mapping row**. Gotcha: Excel/openpyxl erase the text of all but the top-left merged cell, so a title merged down into the mapping row erased that column's name and the table silently started one column later. A test caught this.
- **Widths:** openpyxl stores one `ColumnDimension` per *group* of columns (`min`..`max`) and reports a made-up width of 13 for columns without one. So only entries with `customWidth` are used; the rest get `sheet_format.defaultColWidth` or Excel's 8.43.
- `apply_layout(table, layout)` merges the mapped columns into the browser table: `X*` makes a database column an input column (sheet values ignored, ID column refused), input columns missing from the database are appended, and an unknown non-input column → error listing the available ones. Mapping names refer to *shown* columns: a `#` column cannot be mapped (it stays ignored completely).
- **Data sheet:** `ExcelItemSource(skip_sheet=)` reads the first sheet that is *not* the layout sheet, so the layout sheet may come first. It is only skipped when the layout lives in the database workbook.
- **Settings:** `table_layout_file` (optional, `""` = the item database) and `table_layout_sheet` (default `Word layout`). A missing sheet in the database workbook → automatic table; a separately chosen file must contain the sheet (error otherwise). Layout errors → 500 on `/api/items` and export, prefixed "Word table layout:", so the GUI shows them like database errors. A change to either setting reloads the items.
- `/api/items` adds `layout` (source text, `""` = automatic) and `word_columns`; the GUI shows "Word table: … (n columns)" next to the database name, with the column list as a tooltip.
- **docgen:** `_build_table(..., layout)` creates `header_rows` rows, sets the column widths on *every* cell **before** merging (python-docx adds up the widths of merged cells), then merges per `HeaderCell` and marks all header rows `tblHeader` (repeated on each page). Without a layout, the same code builds one header row from the column names.
- Example: `tools/make_examples.py` also writes `data/layout_example.xlsx` (merges across and down, an input column only in the layout). It is a separate file, so the example database still shows the automatic table and the existing tests are unchanged. Re-running the script rewrites the other example files with new zip timestamps but the same content; don't commit those. Checked visually: LibreOffice rendering of the exported table.
- README gained a "Word table layout" section now (not at release), so the user can try it.

## 2026-11-19 – Save to… (0.3.0 step 4, reduced)
- Step 4 was reduced with the user to "Save to…" so they can try it out. Opening single `.json` files directly was dropped for now, since the library's Other folder tab already loads from any folder.
- `POST /api/sessions` takes an optional `folder` (absolute, existing, local-only, checked inside the handler because a plain save must keep working for remote clients) and now returns the session plus `path` (`SavedSession`).
- GUI: a "Save to…" button opens the folder dialog at the shared library (else My sessions). Both libraries are listed first under "Go to…" (`openBrowse({shortcuts, selectLabel})`), and the button reads "Save here". Plain Save still writes to My sessions.
- Overwrite question: for Save to…, it is skipped only when the same-named session in the target folder *is* the current session's file (`sessionPath`); for plain Save, the previous rule applies. The header shows "(shared library)" or "(other folder)" for the current session's location. `libraryKind(folder)` maps a folder to its library by comparing paths, ignoring trailing separators.

## 2026-11-19 – Session library browser (0.3.0 step 3)
- **Path-based library API** (all local-only, since they read/delete files by path): `GET /api/library?folder=` (empty = `sessions_dir`; a missing user library is just empty because it is created on first save, while another missing folder → 404) returns `{folder, sessions: [summary + file + path], skipped: [{file, error}]}`. `GET`/`DELETE /api/library/session?path=`. The name-based `/api/sessions` endpoints stay for saving into the user library and for the overwrite check.
- `SessionStore.scan()` reports unreadable files (the GUI lists them as "not listed"), and `list()` is built on it. `delete_file` first parses the file as a session, so a wrong path can never delete an unrelated file. Only `.json` files are accepted.
- **Settings:** `sessions_dir` (now changeable at runtime: the store is created per request, like the item source) and new optional `global_sessions_dir` (`""` in the file / `None` in memory, `_OPTIONAL_PATH_KEYS`). Changed values must be existing absolute folders; `""` unsets the shared library. They are set from the library dialog ("Change folder…"), not from the Settings panel, which avoids two places to edit them.
- **GUI:** "Session library…" button replaces the Saved drop-down. A modal dialog has tabs My sessions / Shared library / Other folder (other = any folder via Browse, not persisted; the browse dialog opens on top of it), plus a search box (all terms over name, configuration, file, database) and a table sortable by Name, Configuration, Items, Modified, File (default newest first, `aria-sort`). Click selects; double-click, Enter or Load loads; arrow keys move. Delete asks first, with a warning in the shared library. Helpers `filterSessions`/`sortSessions` live in `logic.js` (Node-tested). Blank configuration sorts last, the same as in the filter lists.
- **Loading from the shared library or another folder:** the header shows "(from the shared library)", and saving writes into My sessions. The "overwrite?" check therefore fires whenever a same-named session exists in My sessions, unless the current session came from My sessions itself (state `sessionKind`/`sessionPath`).

## 2026-11-19 – Template file in Settings (0.3.0 step 2)
- `template_file` joins `GET`/`PUT /api/settings` and the Settings panel ("Word template" + Browse). Like the output folder, only a *changed* value is validated, so a template that has since gone missing doesn't block saving other settings (export reports it then).
- A changed template is checked on save, not just at export: absolute path, `.docx`, readable by python-docx, and has the `{{DECLARED_LIST}}` paragraph (`docgen.check_template`, which shares `_open_template` with export). `.dotx` is not accepted: python-docx cannot open Word template packages.
- `/api/fs` takes `ext` (`.xlsx` default, `.docx`; allowlist `BROWSE_EXTENSIONS`, others → 400). The GUI dialog's `openBrowse({ext})` shows the extension in its hints. Word lock files `~$*.docx` were already hidden by the `~$` rule.

## 2026-11-19 – Windows test fixes (0.3.0 step 1)
- **Caching:** an HTTP middleware sets `Cache-Control: no-cache` on every response (GUI files and API) unless one is already set. The browser may still cache, but revalidates first; `StaticFiles` sends ETag/Last-Modified, so an unchanged file costs only a 304. Before, browsers cached `app.js`/`logic.js` heuristically and kept the old version after an update.
- **Version:** `dlg.get_version()` reads `pyproject.toml` of the checkout first, then the installed metadata. Gotcha: with the editable install (`pip install -e`), the metadata keeps the version from install time and goes stale after a `git pull`. A `pyproject.toml` of another project is ignored (name must be `dlg`). `GET /api/info` → `{version}`; shown as `vX.Y.Z` in the header (not awaited, cosmetic).
- **File-name field:** typing, including deleting the last character, keeps the user's text. Before, an emptied field was refilled with the suggestion immediately, so the old name snapped back mid-edit. Now the suggestion only returns when the field is *left* empty (blur), on Reset name (disabled while following the suggestion), on session load, or when exporting with an empty field.
- `output/` (default export folder) is git-ignored.
- Checked in headless Chromium with real Backspace key events (scratch script over DevTools protocol).

## 2026-11-19 – 0.3.0 planning decisions (user, not yet implemented)
- **Word table layout (step 5) is drawn in Excel, not in the Word template:** header rows hold free text; merged cells (across and down) become merged cells in Word. The last row in use is the mapping row: each cell names a data-sheet column, data rows follow beneath in that order.
- Default location: sheet "Word layout" inside the database workbook; a "Table layout" setting can point to another `.xlsx` file and sheet.
- A column ending in `*` is an input column in either place (data sheet or mapping row). A mapping-row input column does not have to exist in the database; the browser table shows it as an input column.
- Only mapped columns appear in Word, in mapping order; `#` still hides columns from the browser table. No layout sheet → current automatic table. Unknown column in the mapping row → error listing the available columns.
- Excel column widths in the layout sheet set the relative Word column widths, scaled to the table width.
- **Session library (steps 3–4):** "loading other sessions" means opening session files from any folder. Library = user folder (`sessions_dir`) + global/shared folder (new setting) + any folder via Browse. Merging several sessions into one selection → backlog (user wants to think it through).
- The Template-file setting (step 2) stays, but is no longer a prerequisite for step 5.
- **Table font size (step 6)** is a separate setting (default = template font), not read from the layout sheet's mapping row.

## 2026-11-18 – Exit button (0.2.0 step 10)
- `POST /api/shutdown` (local-only, 202) schedules `on_shutdown` as a FastAPI background task, so the response is sent before the server begins to stop. `create_app(..., on_shutdown=None)` → 503 "use Ctrl+C" (tests/programmatic use).
- The launcher builds `uvicorn.Server(uvicorn.Config(...))` itself instead of `uvicorn.run`, and `on_shutdown` sets `server.should_exit = True`: uvicorn's normal graceful shutdown (finishes open requests, process exits with code 0). Verified end to end: clicking Exit in headless Chromium ends the server process.
- GUI: Exit button top right; confirms (stronger wording with unsaved changes); afterwards hides the app, shows a "stopped" panel, clears the dirty flag (no "leave page?" prompt) and tries `window.close()` (only works for script-opened tabs).
- Gotcha: author CSS `display: flex/grid` on `header`/`main` overrides the `hidden` attribute, so "hidden" elements stayed visible; the check passed because it read the attribute. Added a global `[hidden] { display: none !important; }`, and browser checks now assert the computed `display`.

## 2026-11-18 – Output file name & location (0.2.0 step 9)
- `POST /api/export` now **writes** `<output_dir>/<name>.docx` (folder created) and returns `{path, file_name, count}`; the browser download, `render_to_bytes` and `DOCX_MEDIA_TYPE` are gone. Request adds `file_name`, `configuration_name`, `overwrite`.
- Existing file without `overwrite` → 409; the GUI confirms and resends with `overwrite: true` (`api()` errors now carry `status`).
- File names are **validated, not silently changed** (user sees exactly what is written): Windows-forbidden characters `< > : " / \ | ? *` and control chars → 400 on every platform (a name that works on Linux must work on Windows); reserved device names (`CON`, `NUL`, …; shared `WINDOWS_RESERVED_NAMES` from `sessions.py`) → 400; trailing dots/spaces stripped (Windows drops them anyway); `.docx` added if missing (case-insensitive).
- Atomic write: save to `~<stem>.tmp.docx` next to the target, then `os.replace`. `PermissionError` (target open in Word on Windows) → `OutputLockedError` → 423 "close it in Word"; temp file removed.
- `{{CONFIGURATION_NAME}}` is replaced in body paragraphs, (nested) table cells and all headers/footers that are not linked to the previous section (touching a linked one would create an empty header). Replaced **before** the item table is inserted so item data containing the placeholder text is never substituted. Placeholder inside one run keeps that run's formatting; one split across runs is collapsed into the first run. Text boxes are not searched.
- GUI: File name field next to Export, suggested from configuration → session → `declared_list` (Windows-forbidden characters become `_` in the suggestion). Once the user types a name it sticks; emptying the field resumes suggestions; loading a session resets to the suggestion. Enter in the field exports.
- Test gotcha: `Config()` defaults `output_dir` to the project's `output/`, so every exporting test must set `output_dir=tmp_path / ...`.

## 2026-11-18 – Paths & browsing (0.2.0 step 8)
- Browsers cannot reveal real file paths, so the local server lists folders: `GET /api/fs?path=&files=` (`dlg/filebrowser.py`, stdlib only). Empty path = home; a file path lists its folder (dialog opens at the current file). Roots: Home + drives (`os.listdrives()` on 3.12+, else drive-letter probe) on Windows, Home + `/` elsewhere. Hidden: dot-files, Office lock files `~$*.xlsx`, Windows hidden attribute. Sorting case-insensitive. 404 missing / 403 unreadable.
- `GET`/`PUT /api/settings`: `items_file`, `id_column`, `output_dir`. Everything is validated before anything changes; on a failed config write the in-memory values are rolled back. Saved with `save_config` (rewrites the file; user comments are lost — acceptable, file is tool-generated).
- **ID column added to the GUI settings** (not in the original step): the browser check showed a database with another ID header (`Part No`) is unusable without editing the config file, and the real ID header is not yet known.
- Gotcha: the default output folder does not exist until the first export, and the GUI always sends all fields, so validating an *unchanged* `output_dir` blocked every save. Only changed values are validated now.
- Paths must be absolute (a relative path would resolve against the server's working folder).
- **Local-only:** `/api/fs` and `/api/settings` require the client to be `127.0.0.1`/`::1` (FastAPI dependency `require_local`), so they stay local even with `--host` on a network. Tests use `TestClient(client=("127.0.0.1", …))`.
- The item source is now created per request from the current config (`source()`), so a new database applies without restart. The launcher passes the config path to `create_app(config, config_path)`; with `None` (tests) changes stay in memory.
- GUI: collapsible Settings panel, opens automatically when the database fails to load; `<dialog>` browse dialog (Go to…, Up, editable path + Go/Enter, folders then files, double-click selects a file; folder mode selects the current folder). After a database change, filters on vanished columns are dropped and unknown selected IDs are deselected with a warning.

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
