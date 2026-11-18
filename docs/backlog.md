# Backlog

Source of truth for ideas not in the current release plan. The user keeps a private `my_notes.md`; its decisions are copied here.

## Scheduled
Nothing scheduled; next release to be planned in `PLAN.md`.

## Future ideas
- **Configure view** in the GUI for personal settings, persisted in the per-user config file.
- **Document review mode:** read a Word file, compare its item table with the database (Excel), highlight rows/cells that do not match.
- **Searchable library** of session files (shared re-use).
- Export ordering by selection order instead of database order (optional).
- Swap Excel for another `ItemSource` (e.g. SQLite).
- Cache the Excel read if the database grows large.

## Known issues / checks
- **0.2.0 open checks:** run the tool on Windows (setup per README, Browse dialog with drive letters, export with the file open in Word); open an exported `.docx` in real Word (0.1.0 export was confirmed; 0.2.0 adds `{{CONFIGURATION_NAME}}` and data-driven column widths).
- Real Word template from the user: placeholder `{{DECLARED_LIST}}` in a landscape section.
- Many columns may not fit the landscape width; consider a smaller table font or optional per-column width hints if the real database needs it.
