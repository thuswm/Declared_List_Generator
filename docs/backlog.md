# Backlog

Source of truth for ideas not in the current release plan. The user keeps a private `my_notes.md`; its decisions are copied here.

## Scheduled
See `PLAN.md` (Release 0.2.0).

## Future ideas
- **Configure view** in the GUI for personal settings, persisted in the per-user config file.
- **Document review mode:** read a Word file, compare its item table with the database (Excel), highlight rows/cells that do not match.
- **Searchable library** of session files (shared re-use).
- Export ordering by selection order instead of database order (optional).
- Swap Excel for another `ItemSource` (e.g. SQLite).
- Cache the Excel read if the database grows large.

## Known issues / checks
- Open an exported `.docx` in real Word (only checked in LibreOffice).
- Real Word template from the user: placeholder `{{DECLARED_LIST}}` in a landscape section; adjust `TABLE_COLUMNS`.
- Session-name collisions (`a/b` vs `a_b`) share one file.
