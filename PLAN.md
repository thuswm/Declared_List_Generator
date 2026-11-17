# Declared List Generator – Implementation Plan

## Context
New project (empty repo, only CLAUDE.md). DLG lets a user pick items from a database in a browser table, save the selection as a session file, reload it later, and export a Word document containing the declared list as a table. Per CLAUDE.md, this plan is first copied to `PLAN.md` in the repo root (after approval), and each step is checked off there after it is tested.

## Decisions (from user answers)
- **Item DB:** Excel file (v1), accessed behind a `ItemSource` interface so it can be swapped later (SQLite, etc.).
- **Web stack:** FastAPI + pydantic + uvicorn, with vanilla HTML/JS (no build step, no JS framework). Runs locally, opened in Edge/Firefox/Chrome.
- **Excel reading:** `openpyxl` (read-only; the standard, widely adopted library).
- **Word output:** `python-docx`. Recommended because it is the de facto standard, opens an existing template, and handles sections/orientation.
- **Session files:** plain, indented JSON validated by pydantic, with `schema_version`, name, created/modified timestamps, DB source filename, and a list of selected item IDs.
- **Template:** no template supplied, so I create an example `templates/declared_list_template.docx`: portrait pages (cover/intro), a **landscape section** containing the placeholder paragraph `{{DECLARED_LIST}}`, then portrait closing pages. Export finds the placeholder, inserts the table there (inside the landscape section, so orientation is inherited), and removes the placeholder. Header row repeats on each page.

## Layout
```
PLAN.md
pyproject.toml / requirements.txt   (fastapi, uvicorn, pydantic, openpyxl, python-docx, pytest, httpx)
src/dlg/
  models.py        # pydantic: Item, Session
  sources/base.py  # ItemSource protocol
  sources/excel.py # ExcelItemSource (header row -> Item fields; configurable column map)
  sessions.py      # save/load/list session JSON files
  docgen.py        # template + selected items -> .docx
  app.py           # FastAPI routes + static serving
  static/          # index.html, app.js, style.css
data/items.xlsx           # example item DB (generated)
templates/declared_list_template.docx  # example template (generated)
sessions/                 # saved sessions
tools/make_examples.py    # generates example xlsx + docx template
tests/
```

## Phases (each: implement → test → tick in PLAN.md → note next step)
1. [x] **Scaffold:** repo layout, dependencies, `PLAN.md`, `.gitignore`.
2. [x] **Models + Excel source:** Item/Session models; ExcelItemSource; example `items.xlsx` generator; tests.
3. [x] **Session store:** save/load/list JSON; validation + version check; tests.
4. [x] **Word generation:** example template generator (portrait/landscape/portrait); `docgen` placeholder replacement, table styling, repeat header; tests assert placeholder gone, row count, and that the table's section is landscape while others stay portrait.
5. [x] **API:** `GET /api/items`, `GET/POST /api/sessions`, `GET /api/sessions/{name}`, `POST /api/export` (returns .docx); tests with FastAPI TestClient.
6. [x] **Frontend:** items table with checkboxes, search/filter, select-all-visible, selected count, save/load session, export button.
7. [x] **Polish:** error handling (missing file, bad Excel headers, unknown IDs in a loaded session), README run instructions.

## Assumptions (easy to change later)
- Excel: first sheet, first row = headers, one column is a unique ID; the sample uses columns like ID, Name, Description, Category. Declared-list table columns are configurable in one place.
- Single local user; no authentication.

## Verification
- `pytest` for all phases.
- Run `uvicorn dlg.app:app`, open in browser, select items, save, reload the session, export, then open the .docx (or convert with LibreOffice if available) to confirm landscape table pages and portrait elsewhere.

## Progress Log
- **Phase 1 done** (2026-11-17): layout, `pyproject.toml`, `.gitignore`, `.venv` (Python 3.14.7) with deps installed and imports verified. Use `.venv/bin/python` / `.venv/bin/pytest`.
- **Phase 2 done** (2026-11-17): `models.py` (Item, Session), `sources/base.py` (ItemSource Protocol), `sources/excel.py` (ExcelItemSource, `ItemSourceError`), `tools/make_examples.py` -> `data/items.xlsx` (12 items), 9 tests passing.
  - Design notes: IDs are always strings (numeric cells stringified); rows with blank ID skipped; duplicate IDs and missing ID/Name headers raise `ItemSourceError`; unmapped columns are kept in `Item.extra`; Session timestamps are UTC-aware.
  - Gotcha: `tools/` is not a package, so tests must not import it; tests read the generated `data/items.xlsx` instead (re-run `tools/make_examples.py` if it is deleted).
- **Phase 3 done** (2026-11-17): `sessions.py` with `SessionStore` (save/load/list/delete), `SessionError`/`SessionNotFoundError`, `safe_filename`; 23 tests passing total.
  - Design notes: filename is a sanitised stem of the name (non `[\w .-]` chars -> `_`) but the exact display name is kept inside the JSON; names that sanitise to empty are rejected; re-saving keeps `created`, updates `modified`; duplicate selected IDs are removed (order kept); writes are atomic (tmp + rename); `list()` skips corrupt files; files with a newer `schema_version` than supported are refused.
  - Gotcha: two different names that sanitise to the same stem (e.g. `a/b` and `a_b`) share one file and the later save overwrites the earlier. Acceptable for v1; revisit if it matters.
- **Phase 4 done** (2026-11-17): `docgen.py` (`build_document`, `generate_declared_list`, `render_to_bytes`, `DocgenError`, `TABLE_COLUMNS`), `tools/make_examples.py` now also writes `templates/declared_list_template.docx` (A4 portrait intro / landscape section with `{{DECLARED_LIST}}` / portrait notes). 36 tests passing total.
  - Verified in LibreOffice (40 items -> 7 pages): portrait, 5 landscape table pages with repeating header, portrait.
  - Design notes: the placeholder paragraph must contain exactly `{{DECLARED_LIST}}` (matched on joined text, so split runs are fine); table column layout lives in `TABLE_COLUMNS` (header, value fn, relative weight); empty item list raises `DocgenError`; items are listed in the order passed in (the API decides DB order vs selection order); control characters are stripped from cell text; template file is never modified.
  - Gotchas found: (1) if the placeholder paragraph holds the section break (Word does this when the break directly follows it) it must not be deleted and the table must be inserted *before* it, else the table lands in the next section; (2) LibreOffice uses the `tblGrid` column widths, so set column widths as well as cell widths; (3) `w:tblPr` children must stay in schema order (edit `tblW` in place, use `insert_element_before` for borders) or Word may report a corrupt file. Not yet opened in real Word - worth a manual check.
- **Phase 5 done** (2026-11-17): `app.py` with `create_app(settings)` factory + module-level `app`; `Settings` (env overrides `DLG_ITEMS`, `DLG_TEMPLATE`, `DLG_SESSIONS`); 44 tests passing; smoke-tested with real uvicorn + curl.
  - Routes: `GET /api/items` -> `{source, items}`; `GET /api/sessions` -> summaries (name, modified, source, count); `GET /api/sessions/{name}`; `POST /api/sessions` `{name, selected_ids}` (create/overwrite, returns saved Session); `DELETE /api/sessions/{name}` (204; added beyond the original plan); `POST /api/export` `{selected_ids}` -> .docx download (`declared_list.docx`). Static GUI is mounted at `/` from `src/dlg/static/` after the API routes.
  - Design notes: Excel is re-read on every request (edits show up without restart; fine for small DBs, cache if it gets large); export rows follow **database order**, not click order; export rejects empty selection (400) and unknown IDs (422, lists them); save is lenient and does not check IDs against the DB; missing DB/template -> 500 with message; missing session 404, bad name 400, corrupt session file 422. Endpoints are sync `def` so file I/O runs in FastAPI's threadpool.
  - Gotcha: Starlette's TestClient prints a deprecation warning ("install `httpx2` instead"). Harmless; left as is because httpx2 is a niche package (ask before adding).
- **Phase 6 done** (2026-11-17): GUI in `src/dlg/static/` (`index.html`, `style.css`, `app.js`, `logic.js`). 47 pytest tests passing (includes the 9 Node tests in `tests/js/logic.test.js` via `tests/test_frontend.py`, skipped if Node is absent).
  - Features: searchable/filterable item table (multi-term AND search across all fields incl. extra columns), category filter, "selected only", header select-all-visible checkbox (tri-state), click anywhere on a row to toggle, selected/shown counts, save (empty name rejected, overwrite confirmation), load, delete (confirm), export (.docx download), unsaved-changes confirm + beforeunload warning, light/dark theme.
  - Design notes: pure logic lives in `logic.js` (ES module, no DOM) so Node can test it; `app.js` only uses `textContent` for data (no innerHTML -> spreadsheet content cannot inject markup); loading a session drops IDs no longer in the DB and shows a warning listing them (they are not kept, so re-saving removes them); export button/Clear disabled when nothing is selected. `extra` Excel columns are searchable but not displayed.
  - Verified end-to-end in headless Firefox (22 checks via WebDriver BiDi: search, select-visible, tri-state header, row click, selected-only, save/load/delete, unknown-ID warning, export, markup-in-name stays text) with no console errors. That driver was a throwaway script (not in the repo; needs Firefox) - no browser test framework was added, per the dependency rule.
  - Bug found by the browser test and fixed in `app.py`: session names containing `/` (sent as `%2F`) returned 404 on load/delete because `{name}` cannot match a slash; routes now use `{name:path}` (regression test added).
  - Gotchas: `firefox --headless` refuses to start if the user's own Firefox is running - use `--no-remote --profile <tmpdir>`; never `pkill -f` a pattern that appears in your own command line.
- **Phase 7 done** (2026-11-17): README.md; `python -m dlg` launcher (`src/dlg/__main__.py`: `--host` default 127.0.0.1, `--port`, `--no-browser`); 49 tests passing.
  - Audit against CLAUDE.md: every module/class/function has a docstring; added the type hints that were missing on a few helpers.
  - Missing/invalid item DB: already surfaces in the GUI as an error banner (role=alert) with the specific message, Export disabled; verified in headless Firefox. No code change needed.
  - Decision: unknown IDs in a loaded session are **dropped with a warning** (not kept). Revisit if sessions need to survive temporary DB changes.
  - Default bind is localhost because there is no authentication.

## Post-completion fixes
- (2026-11-17) User saw a 404 for the favicon request in the server log. Fixed by declaring an inline SVG icon in `index.html` (`<link rel="icon" href="data:...">`); verified no favicon request reaches the server in Firefox. Regression test in `tests/test_frontend.py`.
- (2026-11-17) `python -m dlg` failed for the user because system Python was used instead of `.venv`; README now explains the venv.

## Status: all planned phases complete
Nothing is committed yet (waiting for the user's go-ahead).

### Open items / ideas for later
- Open an exported .docx in real Word (only checked in LibreOffice so far).
- Real template from the user: placeholder must be `{{DECLARED_LIST}}` in a landscape section; adjust `TABLE_COLUMNS` to the needed columns.
- Session-name collisions (`a/b` vs `a_b`) share a file.
- Excel is re-read per request; add caching if the database grows large.
- Future: swap Excel for another `ItemSource`; a searchable library of session files; ordering by selection instead of database order if wanted.
