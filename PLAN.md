# Declared List Generator – Plan

Keep this file short (about one screen). Details live elsewhere:
[decisions](docs/decisions.md) · [backlog](docs/backlog.md) · [changelog](CHANGELOG.md) · [archive](docs/archive/) · [README](README.md)

## Status
- **Released:** `v0.1.0` (2026-11-17) – first workable version. Full history: [docs/archive/release-0.1.0-plan.md](docs/archive/release-0.1.0-plan.md).
- **Current:** Release 0.2.0 in progress on branch `release/0.2.0` (steps 1–7 of 11 done). Plan revised 2026-11-18 after the Windows test (approved).
- **Next step:** step 8 – paths & browsing (file/folder browser to set item DB, output folder; persist to config).

## Release 0.2.0 – plan
Design decisions for this release: [docs/decisions.md](docs/decisions.md) (sections "0.2.0 …").

### Steps (each: implement → test → tick here → log decisions)
1. [x] **Config file (done 2026-11-17, 61 tests pass):** platform path resolver, load/create-with-defaults, writer, `--config`; env vars removed.
2. [x] **README for Linux and Windows (done 2026-11-18):** setup/run commands for both (PowerShell + cmd: `.venv\Scripts\python`, quoting of `.[dev]`), config file location on both.
3. [x] **Columns from the Excel file (done 2026-11-18, 71 Python + 9 JS tests pass):** `id_column` setting (default `ID`); item = ID + ordered column values; headers starting with `#` ignored everywhere; API returns the column list; GUI table and Word table built from it (Word: exactly the non-omitted Excel columns, no extra `No.` column; widths from content length); remove hard-coded `TABLE_COLUMNS`, `Item.name/description/category` and the category dropdown; v1 sessions still load (they store IDs only); tests.
4. [x] **Excel-style column filters (done 2026-11-18, 71 Python + 15 JS tests pass, checked in headless Chromium):** filter button per column header → list of distinct values with checkboxes + text box; filters combine (AND) with the global search and "only selected"; active filters visible and clearable; logic in `logic.js` with Node tests.
5. [x] **Session schema v2 (done 2026-11-18, 84 Python + 15 JS tests pass):** `configuration_name`, `user_values`, v1 → v2 migration on load, name-collision fix; tests.
6. [x] **Configuration name in GUI (done 2026-11-18, 84 Python + 17 JS tests pass, checked in headless Chromium):** field when saving; shown in the session load list and in the header after loading.
7. [x] **User-input columns (done 2026-11-18, 88 Python + 20 JS tests pass, checked in headless Chromium):** header ending in `*`; editable cells in the table, stored in session, exported as normal columns; tests incl. Node tests.
8. [ ] **Paths & browsing:** file/folder browser endpoint + GUI dialog; set item DB location and output folder; persist to config.
9. [ ] **Output file name & location:** name field, write to `output_dir`, overwrite confirmation, show saved path; `{{CONFIGURATION_NAME}}` placeholder.
10. [ ] **Exit button:** shutdown endpoint + GUI button; clean server stop; test.
11. [ ] **Release:** test run on Windows, re-check export in real Word, README + CHANGELOG, bump to 0.2.0, tag `v0.2.0`, trim this file, archive the plan.

