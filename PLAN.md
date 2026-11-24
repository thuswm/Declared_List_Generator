# Declared List Generator – Plan

Keep this file short (about one screen). Details live elsewhere:
[decisions](docs/decisions.md) · [backlog](docs/backlog.md) · [changelog](CHANGELOG.md) · [archive](docs/archive/) · [README](README.md)

## Status
- **Released:** `v0.4.0` (2026-11-20, GitHub pre-release) – blank Word columns + layout warnings, `-` placeholder, lenient column names, "TBD" for empty input cells, export summary, Check Word table dialog, ISO dates, `start_dlg.bat`. Plan: [docs/archive/release-0.4.0-plan.md](docs/archive/release-0.4.0-plan.md).
- Earlier: `v0.3.0` (2026-11-19) – [plan](docs/archive/release-0.3.0-plan.md); `v0.2.0`, `v0.1.0` – [archive](docs/archive/).
- **Open checks:** user tries 0.4.0 (and the 0.3.0 open checks) on Windows. See [backlog](docs/backlog.md).
- **Next step:** 0.5.0 step 3 (Open buttons in Settings).

## Release 0.5.0 – plan (approved 2026-11-24)
From the user's 0.4.0 feedback (`my_notes.md`, 2026-11-24). Each step: implement → test → tick here → log decisions in `docs/decisions.md`.
1. [x] **`\blank` marker for blank columns:** `\blank` (any case, spaces around ignored) in a mapping cell = blank Word column on purpose; it becomes the documented marker (README, example file, messages). Reason: a lone `-` is easy to get wrong (autocorrect turns it into `–`, which is read as an unknown name → warning, and twice → "mapped twice" error; a plain `-` works in tests, so that is the likely cause of the 0.4.0 reports). `-` and its look-alikes (`–`, `—`, `−`, `‐`) keep working as blank so 0.4.0 layout files still load. `\blank*` → error. Unknown/duplicate-name messages quote the cell text. Not confirmed with the user's file (not at hand).
2. [x] **ID column as a pick list:** Settings → *ID column* gets a drop-down of the headers of the item database in the *Item database* field (new `GET /api/headers?items_file=…`, re-read when that field changes). Free typing stays possible (input + datalist) if the file cannot be read.
3. [ ] **Open buttons in Settings:** *Open* next to Item database, Word template, Table layout and Output folder → opens the file in Excel/Word (folder in Explorer). New `POST /api/open-setting {setting}`: only opens the *saved* path of that setting (no arbitrary paths from the browser); unsaved edits → "Save settings first". The layout is re-read on every load/export, so after editing in Excel: save in Excel → *Check Word table* / export shows the change.
4. [ ] **Rotated header text:** read from the Excel cell format (Format Cells → Alignment → Orientation): text rotated down (−90°, starts at the top) → Word `tbRl`; rotated up (+90°) → `btLr`. Header row height: the Excel row height if set by hand, else estimated from the longest rotated text and the font size ("at least" height). Other angles → shown horizontal + warning. *Check Word table* dialog shows the rotation.
5. [ ] **Column widths (question, no code):** yes – the Excel column widths are relative weights, scaled to the usable page width (page width − margins) of the template section holding the placeholder. Add this to README and show % per column in *Check Word table*.
6. [ ] **Release:** README + CHANGELOG, version 0.5.0, tag `v0.5.0` (pre-release), archive plan; Windows check by user.
