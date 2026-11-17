# Decisions & Gotchas

Architecture decision log, newest first. Full per-phase history of release 0.1.0:
[archive/release-0.1.0-plan.md](archive/release-0.1.0-plan.md).

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
