# Decisions & Gotchas

Architecture decision log, newest first. Full per-phase history of release 0.1.0:
[archive/release-0.1.0-plan.md](archive/release-0.1.0-plan.md).

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
- **Sessions:** JSON validated by pydantic, atomic writes, newer `schema_version` refused. Known gotcha: names that sanitise to the same stem (`a/b`, `a_b`) share a file.
- **Word export:** placeholder paragraph must be exactly `{{DECLARED_LIST}}` in a landscape section; columns defined in `TABLE_COLUMNS`; rows follow database order. Gotchas: keep the section break when the placeholder paragraph holds it; set `tblGrid` widths (LibreOffice); keep `w:tblPr` child order (Word). Not yet verified in real Word.
- **API:** Excel re-read on every request; export rejects empty (400) / unknown IDs (422). Session routes use `{name:path}`.
- **GUI:** pure logic in `logic.js` (Node-testable); data written with `textContent` only (no `innerHTML`). Unknown IDs in a loaded session are dropped with a warning.
- **Security:** default bind is 127.0.0.1 because there is no authentication.
