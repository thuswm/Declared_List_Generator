# Backlog

Source of truth for ideas not in the current release plan. The user keeps a private `my_notes.md`; its decisions are copied here.

## Scheduled
Release 0.3.0 (see `PLAN.md`): Windows test fixes (page caching, file-name field), template file in Settings, session browser, open session files from anywhere, multi-level Word table header, table font size, open document after export.

## Future ideas
- **Document review mode:** read a Word file, compare its item table with the database (Excel), highlight rows/cells that do not match.
- **Searchable library** of session files (shared re-use): user + global library browser scheduled for 0.3.0.
- **Merge several sessions** into one selection (user finds it interesting, wants to think it through first: conflicts in user values, configuration name, order).
- Swap Excel for another `ItemSource` (e.g. SQLite).
- Cache the Excel read if the database grows large.

## Known issues / checks
- **0.2.0 Windows test (done 2026-11-19 by the user):** works in general. Findings → 0.3.0 step 1: page needed F5 several times to load the new version; file-name field snapped back to the old name when the last character was deleted; "Saved" drop-down is not searchable (→ session browser); Word table needs a three-level header (→ step 5).
- Real Word template from the user: placeholder `{{DECLARED_LIST}}` in a landscape section.
- Many columns may not fit the landscape width: table font size scheduled for 0.3.0; per-column width hints only if still needed.
