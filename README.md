# Declared List Generator (DLG)

Pick items from a database in your browser, save the selection as a session,
and export a Word document containing the declared list as a table.

## Setup

Requires Python 3.11+. The project's dependencies live in a virtual
environment (`.venv`), so a plain `python -m dlg` from the system Python fails
with `No module named dlg`. Either activate the venv or call its Python directly.

```bash
python -m venv .venv                       # once
.venv/bin/pip install -e '.[dev]'          # once; drop '[dev]' if you don't need the tests
.venv/bin/python tools/make_examples.py    # once; writes data/items.xlsx and the example template
```

## Run

Run from the project folder, using the venv's Python:

```bash
.venv/bin/python -m dlg                    # starts the server and opens the browser
```

or activate the venv once per terminal and then use plain `python`
(bash/zsh; on Windows use `.venv\Scripts\activate`):

```bash
source .venv/bin/activate
python -m dlg
```

Options: `--config`, `--port 9000`, `--no-browser`, `--host`. The server listens on
`127.0.0.1` only by default. There is no authentication, so do not expose it
on an untrusted network. If the port is already in use, pick another with
`--port`. Stop the server with Ctrl+C.

### Configuration

Each user has a personal TOML file, created with defaults the first time the
tool starts (or whenever it is missing; delete it to reset):

| Platform      | Location                                                   |
|---------------|------------------------------------------------------------|
| Linux / macOS | `$XDG_CONFIG_HOME/dlg/config.toml`, else `~/.config/dlg/config.toml` |
| Windows       | `%APPDATA%\dlg\config.toml`                                |

Use a different file with `python -m dlg --config path/to/config.toml`.

| Key             | Default                                    | Meaning                                  |
|-----------------|--------------------------------------------|------------------------------------------|
| `items_file`    | `data/items.xlsx` (in the project)         | Excel item database                      |
| `template_file` | `templates/declared_list_template.docx`    | Word template                            |
| `sessions_dir`  | `sessions/` (in the project)               | Folder for session files                 |
| `output_dir`    | `output/` (in the project)                 | Folder for generated Word documents (used from a later step) |
| `host`, `port`  | `127.0.0.1`, `8000`                        | Where the server listens                 |

Relative paths in the file are resolved against the folder containing it.
`--host` / `--port` on the command line override the file. The old `DLG_*`
environment variables are no longer used.

## Using it

1. Tick the items that belong on the declared list (search, filter by category,
   or click a row). The header checkbox selects everything currently shown.
2. Enter a name and **Save** to store the selection as a session; **Load** a
   saved session later.
3. **Export Word document** downloads `declared_list.docx`.

The Excel file is re-read on every page load, so you can edit it and refresh.

## File formats

**Item database (Excel).** First sheet, first row = headers. Required columns:
`ID` (unique) and `Name`. Optional: `Description`, `Category`. Any other
column is kept and is searchable, but not shown. Rows without an ID are
skipped; duplicate IDs are an error.

**Session file (JSON).** One readable file per session in `sessions/`:

```json
{
  "schema_version": 1,
  "name": "Project A",
  "created": "2026-11-17T18:34:41Z",
  "modified": "2026-11-17T18:34:41Z",
  "source": "items.xlsx",
  "selected_ids": ["DL-001", "DL-003"]
}
```

Session files from a newer schema version are refused with a clear message.
The file name is a sanitised form of the session name.

**Word template.** Any `.docx`. Put a paragraph containing exactly
`{{DECLARED_LIST}}` where the table should go. The table is inserted there and
inherits that section's page setup, so put the placeholder in a **landscape
section** (with portrait sections before and after, as in the example
template). The table has a repeating header row. Change the columns in
`TABLE_COLUMNS` in `src/dlg/docgen.py`.

## Behaviour worth knowing

- Export rows follow **database order**, not the order you clicked.
- Loading a session whose items are no longer in the database skips them and
  lists their IDs in a warning. They are dropped if you save again.
- Session names that differ only by characters not allowed in file names
  (for example `a/b` and `a_b`) share one file; the later save overwrites.

## Tests

```bash
.venv/bin/pytest          # Python tests (also runs the JS tests if Node is installed)
node --test tests/js      # JS tests alone
```

## Layout

```
src/dlg/
  config.py        per-user config file
  models.py        Item and Session models
  sources/         ItemSource interface and the Excel implementation
  sessions.py      session JSON storage
  docgen.py        Word generation
  app.py           FastAPI routes + static GUI
  static/          index.html, app.js, logic.js, style.css
tools/make_examples.py   example database and template
PLAN.md                  design decisions and progress log
```

Swapping the item database later means writing another class with a
`name` and `load_items()` (see `sources/base.py`) and using it in `app.py`.
