# Declared List Generator (DLG)

Pick items from a database in your browser, save the selection as a session,
and export a Word document containing the declared list as a table.

## Setup

Requires Python 3.11 or newer. The project's dependencies are installed into a
virtual environment (`.venv`) in the project folder, so a plain `python -m dlg`
from the system Python fails with `No module named dlg`. Either activate the
venv or call its Python directly (both shown below).

All commands are run from the project folder. Each setup step is needed only once.

### Linux / macOS (bash, zsh)

```bash
python3 -m venv .venv
.venv/bin/pip install -e '.[dev]'          # drop '[dev]' if you don't need the tests
.venv/bin/python tools/make_examples.py    # writes data/items.xlsx and the example template
```

### Windows (PowerShell or Command Prompt)

Install Python from [python.org](https://www.python.org/downloads/windows/)
(it includes the `py` launcher). Then:

```bat
py -3 -m venv .venv
.venv\Scripts\pip install -e ".[dev]"
.venv\Scripts\python tools\make_examples.py
```

Use double quotes around `".[dev]"`: Command Prompt does not understand single quotes.

## Run

Linux / macOS:

```bash
.venv/bin/python -m dlg                    # starts the server and opens the browser
```

Windows:

```bat
.venv\Scripts\python -m dlg
```

Alternatively, activate the venv once per terminal and then use plain `python -m dlg`:

| Shell              | Activate                         |
|--------------------|----------------------------------|
| bash / zsh         | `source .venv/bin/activate`      |
| PowerShell         | `.venv\Scripts\Activate.ps1`     |
| Command Prompt     | `.venv\Scripts\activate.bat`     |

If PowerShell refuses to run `Activate.ps1` ("running scripts is disabled on this
system"), either call `.venv\Scripts\python` directly as above, or allow local
scripts once for your user:
`Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.

Options: `--config`, `--port 9000`, `--no-browser`, `--host`. The server listens on
`127.0.0.1` only by default. There is no authentication, so do not expose it
on an untrusted network. If the port is already in use, pick another with
`--port`. Stop the server with Ctrl+C in the terminal.

### Configuration

Each user has a personal TOML file, created with defaults the first time the
tool starts (or whenever it is missing; delete it to reset):

| Platform      | Location                                                   |
|---------------|------------------------------------------------------------|
| Linux / macOS | `$XDG_CONFIG_HOME/dlg/config.toml`, else `~/.config/dlg/config.toml` |
| Windows       | `%APPDATA%\dlg\config.toml`                                |

Use a different file with `python -m dlg --config path/to/config.toml`.

To open it for editing: `xdg-open ~/.config/dlg/config.toml` (Linux),
`notepad "$env:APPDATA\dlg\config.toml"` (PowerShell) or
`notepad "%APPDATA%\dlg\config.toml"` (Command Prompt).

| Key             | Default                                    | Meaning                                  |
|-----------------|--------------------------------------------|------------------------------------------|
| `items_file`    | `data/items.xlsx` (in the project)         | Excel item database                      |
| `template_file` | `templates/declared_list_template.docx`    | Word template                            |
| `sessions_dir`  | `sessions/` (in the project)               | Folder for session files                 |
| `output_dir`    | `output/` (in the project)                 | Folder for generated Word documents (used from a later step) |
| `id_column`     | `ID`                                       | Header of the item-database column holding the unique item IDs |
| `host`, `port`  | `127.0.0.1`, `8000`                        | Where the server listens                 |

Relative paths in the file are resolved against the folder containing it.
Windows paths: inside double quotes a backslash must be doubled
(`"C:\\Data\\items.xlsx"`); single quotes (`'C:\Data\items.xlsx'`) or forward
slashes (`"C:/Data/items.xlsx"`) also work.
`--host` / `--port` on the command line override the file. The old `DLG_*`
environment variables are no longer used. Settings added by a newer version
are appended to an existing file with their defaults, so you can see and change them.

## Using it

1. Tick the items that belong on the declared list (click the checkbox or the
   row). The header checkbox selects everything currently shown.
   - **Search** looks in all shown columns (every word must match).
   - **Column filters** work like Excel's: click ▾ in a column header, tick the
     values to show (type in the box to narrow the list; Enter = OK). The list
     only offers values left by the other columns' filters. A filtered column's
     ▾ is highlighted; **Clear filters** removes them all.
   - Search, column filters and **Selected only** all apply together.
     Filters are not saved in the session.
2. Enter a name and **Save** to store the selection as a session; **Load** a
   saved session later.
3. **Export Word document** downloads `declared_list.docx`.

The Excel file is re-read on every page load, so you can edit it and refresh.

## File formats

**Item database (Excel).** First sheet, first row = headers. The columns of
the sheet are the columns of the tool: every column with a header is shown in
the browser table and written to the Word table, in sheet order.

- **ID column:** one column holds a unique ID per item (selections are saved
  by ID). Its header is set by `id_column` in the configuration file (default
  `ID`). Rows without an ID are skipped; duplicate IDs are an error.
- **Omitting a column:** start its header with `#` (e.g. `#Internal note`).
  The tool then ignores it completely: not shown, not searched, not exported.
  Columns without a header are ignored too.
- Two shown columns with the same header are an error.
- Dates are shown as `YYYY-MM-DD`; whole numbers without a decimal part.

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
template). The table has a repeating header row and exactly the shown
(non-`#`) columns of the item database; column widths follow the length of
their content.

## Behaviour worth knowing

- Export rows follow **database order**, not the order you clicked.
- Loading a session whose items are no longer in the database skips them and
  lists their IDs in a warning. They are dropped if you save again.
- Session names that differ only by characters not allowed in file names
  (for example `a/b` and `a_b`) share one file; the later save overwrites.

## Tests

```bash
.venv/bin/pytest          # Linux/macOS: Python tests (also runs the JS tests if Node is installed)
.venv\Scripts\pytest      # Windows
node --test tests/js      # JS tests alone (needs Node.js)
```

## Layout

```
src/dlg/
  config.py        per-user config file
  models.py        Column, Item, ItemTable and Session models
  sources/         ItemSource interface and the Excel implementation
  sessions.py      session JSON storage
  docgen.py        Word generation
  app.py           FastAPI routes + static GUI
  static/          index.html, app.js, logic.js, style.css
tools/make_examples.py   example database and template
PLAN.md                  design decisions and progress log
```

Swapping the item database later means writing another class with a
`name` and `load()` returning an `ItemTable` (see `sources/base.py`) and using it in `app.py`.
