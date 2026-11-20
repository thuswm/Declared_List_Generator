# Declared List Generator – Plan

Keep this file short (about one screen). Details live elsewhere:
[decisions](docs/decisions.md) · [backlog](docs/backlog.md) · [changelog](CHANGELOG.md) · [archive](docs/archive/) · [README](README.md)

## Status
- **Released:** `v0.3.0` (2026-11-19, GitHub pre-release); Windows feedback received 2026-11-20 → release 0.4.0 below. Earlier plans: [archive](docs/archive/).
- **Next step:** step 5 (export summary of empty input cells). Work on branch `release/0.4.0`.

## Release 0.4.0 – plan (approved 2026-11-20)
Steps (each: implement → test → tick here → log decisions in `docs/decisions.md`). Planning decisions of 2026-11-20: [decisions](docs/decisions.md).
1. [x] **Blank Word columns instead of layout errors:** in the layout sheet's mapping row, an *empty cell* inside the table, or a *name the database lacks* (without `*`), no longer stops loading → column exported empty (header from the rows above) + **warning**. A **`-`** in a mapping cell means "blank on purpose": exported empty, **no warning**, also usable at the table edge (span stays first → last filled mapping cell). `parse_layout`/`apply_layout` collect warnings; `/api/items` and `/api/export` return `warnings: [...]`; GUI shows a `warn` status and a ⚠ on the "Word table: …" label (tooltip lists them). Blank columns are Word-only. Still errors: duplicate names, a `*` with no name, merges into the mapping row, ID column as input.
2. [x] **Unmapped database columns:** already allowed; add an explicit test and a README sentence.
3. [x] **O2 Lenient column matching:** mapping names match database columns ignoring case and repeated/leading/trailing spaces; a non-exact match gives a warning ("'part no' matched 'Part No'"). Ambiguous match (two DB columns) → error.
4. [x] **"TBD" for empty input cells:** settings `mark_empty_inputs` (bool, default off) and `empty_input_text` (default `TBD`). When on, empty **input (`*`) column** cells are written as that text in Word (layout and automatic table). Empty database values and blank columns stay empty. Settings panel: checkbox + text field.
5. [ ] **O4 Export summary:** export response lists empty input cells per column ("Qty: 3 empty → TBD"); GUI shows it in the export status.
6. [ ] **ISO dates in the GUI:** session library "Modified" and `sessionLabel` as `YYYY-MM-DD HH:MM` (local time). Helper in `logic.js` + Node test.
7. [ ] **O1 Word-table preview/check:** "Check Word table" view in the browser: header grid (merges, widths as proportions, blank columns marked) + warnings, from a new `GET /api/layout`.
8. [ ] **O3 Windows start script:** `start_dlg.bat` (activates `.venv`, runs `python -m dlg`); README section.
9. [ ] **Release:** README + CHANGELOG, version 0.4.0, tag `v0.4.0` (GitHub pre-release), archive plan; Windows check by user (gaps/`-`/unknown names, lenient matching, TBD + summary, dates, preview, .bat; plus the 0.3.0 open checks).
