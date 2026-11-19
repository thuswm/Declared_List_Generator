# Declared List Generator – Release 0.3.0 plan (archived)

Archived at release 2026-11-19. Links are relative to the repo root:
[decisions](docs/decisions.md) · [backlog](docs/backlog.md) · [changelog](CHANGELOG.md) · [archive](docs/archive/) · [README](README.md)

## Release 0.3.0 – plan (approved 2026-11-19)
Steps (each: implement → test → tick here → log decisions in `docs/decisions.md`). Design decisions of 2026-11-19: [decisions](docs/decisions.md).
1. [x] **Windows test fixes:** `Cache-Control: no-cache` on the GUI files (no F5 needed after an update) and app version in the header; file-name field no longer snaps back to the old name when emptied (suggestion only refilled when the field is left empty, plus a "reset to suggestion" button); `output/` git-ignored.
2. [x] **Template file in Settings:** choose the Word template with Browse like the other paths.
3. [x] **Session library browser:** table of all sessions in a folder (name, configuration, items, modified, file) with search and sorting. Folder switch: user library (`sessions_dir`), global library (new setting `global_sessions_dir`), or any folder picked with Browse. Load/delete from there; replaces the "Saved" drop-down.
4. [x] **Save to…** (reduced 2026-11-19, user to try it out): save the current session into a chosen folder (shared library, My sessions or any folder) via the folder dialog; asks before overwriting there. Plain Save still writes to My sessions. Opening single `.json` files directly: dropped for now (Other folder tab covers it).
5. [x] **Multi-level Word table header from an Excel layout sheet:** header rows with free text and merged cells (across/down) rebuilt as merged cells in Word; last used row = mapping row (data-sheet column names, data rows beneath in that order). Default sheet "Word layout" in the database workbook; "Table layout" setting may point to another `.xlsx` + sheet. `*` marks input columns in either place (mapping-row input columns need not exist in the database; the browser table shows them). Only mapped columns go to Word; `#` still hides columns in the browser. Excel column widths → relative Word widths, scaled to the table width. Unknown column in the mapping row → error listing available columns. No layout sheet → current automatic table.
6. [x] **Table font size:** separate setting for the Word table font size (default = template font); not taken from the layout sheet. With a size set, cells get compact spacing.
7. [x] **Open the document after export:** "Open document" / "Open folder" buttons after a successful export; local-only endpoint, only files in the output folder (`os.startfile` on Windows, `xdg-open` on Linux).
8. [x] **Release (2026-11-19):** README + CHANGELOG, version 0.3.0, tag `v0.3.0` (GitHub pre-release), plan archived. **Not done:** user's try-out of steps 3–7 on Windows and Open document/folder with real Word/Explorer – carried over as open checks in `docs/backlog.md`.
