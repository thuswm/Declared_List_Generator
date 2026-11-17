# Declared List Generator – Plan

Keep this file short (about one screen). Details live elsewhere:
[decisions](docs/decisions.md) · [backlog](docs/backlog.md) · [changelog](CHANGELOG.md) · [archive](docs/archive/) · [README](README.md)

## Status
- **Released:** `v0.1.0` (2026-11-17) – first workable version. Full history: [docs/archive/release-0.1.0-plan.md](docs/archive/release-0.1.0-plan.md).
- **Current:** Release 0.2.0 in progress on branch `release/0.2.0` (step 1 of 8 done).
- **Next step:** step 2 – session schema v2 (`configuration_name`, `user_values`, v1 migration, name-collision fix).

## Release 0.2.0 – plan (final)

### Design decisions proposed
- **Config file:** TOML (read with stdlib `tomllib`; a tiny hand-written writer for the flat settings, no new dependency). One file per user, created with defaults on first run or if missing. Location by platform, resolved with stdlib only:
  - Linux/macOS: `$XDG_CONFIG_HOME/dlg/config.toml`, default `~/.config/dlg/config.toml`
  - Windows: `%APPDATA%\dlg\config.toml`
  - Override for testing: `--config <path>` on the launcher.
  - Replaces the `DLG_*` environment variables (removed).
- **Settings:** `items_file`, `template_file`, `sessions_dir`, `output_dir`, (`host`, `port`).
- **Folder/file browsing:** server-side browser dialog in the GUI (the server runs locally, so it can list the user's folders; works the same on Linux and Windows, no `tkinter`). Restricted to localhost.
- **Output:** server writes the .docx to `output_dir` under a user-given file name (default proposed from the configuration name), and the GUI reports the saved path.
- **User-input columns:** marked in the Excel header row with a trailing `*` (e.g. `Serial number*`). Per item, the user types a value in the table; values are saved in the session and shown in the Word table as extra column(s).
- **Session schema v2:** adds `configuration_name` and `user_values` (`{item_id: {column: value}}`). v1 files still load (empty defaults). Session filename collision fix (`a/b` vs `a_b`) handled in the same change.
- **Exit button:** `POST /api/shutdown`, accepted from localhost only; GUI shows "server stopped" afterwards.

### Steps (each: implement → test → tick here → log decisions)
1. [x] **Config file (done 2026-11-17, 61 tests pass):** platform path resolver, load/create-with-defaults, writer, `--config`; remove env vars; update README; tests.
2. [ ] **Session schema v2:** `configuration_name`, `user_values`, v1 → v2 migration on load, name-collision fix; tests.
3. [ ] **Configuration name in GUI:** field when saving; shown in the session load list and in the header after loading.
4. [ ] **User-input columns:** marker parsing in `ExcelItemSource`, API exposes them, editable cells in the table, stored in session, included in export (`TABLE_COLUMNS`); tests incl. Node logic tests.
5. [ ] **Paths & browsing:** file/folder browser endpoint + GUI dialog; set item DB location and output folder; persist to config.
6. [ ] **Output file name & location:** name field, write to `output_dir`, overwrite confirmation, show saved path.
7. [ ] **Exit button:** shutdown endpoint + GUI button; clean server stop; test.
8. [ ] **Release:** manual check in real Word (if available), README + CHANGELOG, bump version to 0.2.0, tag `v0.2.0`, trim this file and archive the plan.

### Decisions (2026-11-17, all open questions answered)
- User-input columns: header ends with `*` (e.g. `Serial number*`). Column names come from the Excel sheet; the `*` is stripped for display and export.
- Values typed by the user per selected item are saved in the session file **and** written to the Word table as extra columns (one per `*` column, automatic).
- Output: written to the configured output folder only (no browser download in this version); name set by the user.
- Configuration name also appears in the exported document via a `{{CONFIGURATION_NAME}}` placeholder (add to example template; handled in step 6).
