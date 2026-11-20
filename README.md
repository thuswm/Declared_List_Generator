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
.venv/bin/python tools/make_examples.py    # writes the example database, table layout and template
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

Or **double-click `start_dlg.bat`** in the project folder: no console commands
needed. For a desktop icon, right-click `start_dlg.bat` → *Show more options* →
*Send to* → *Desktop (create shortcut)*. A console window stays open while the
tool runs; closing it (or **Exit** in the browser) stops the tool. If `.venv`
is missing or the tool stops with an error, the window stays open so the
message can be read. Arguments are passed on, e.g. `start_dlg.bat --port 8001`.

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
`--port`. Stop the tool with the **Exit** button at the top right of the page
(it warns about unsaved changes first), or with Ctrl+C in the terminal.

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
| `sessions_dir`  | `sessions/` (in the project)               | Folder for session files ("My sessions") |
| `global_sessions_dir` | empty                                | Shared session library, e.g. on a network drive ("Shared library") |
| `output_dir`    | `output/` (in the project)                 | Folder for generated Word documents      |
| `id_column`     | `ID`                                       | Header of the item-database column holding the unique item IDs |
| `table_layout_file` | empty                                  | Workbook with the Word table layout; empty = the layout sheet inside the item database |
| `table_layout_sheet` | `Word layout`                         | Name of the layout sheet                 |
| `table_font_size` | `0`                                      | Word table font size in pt (4–72, steps of 0.5); 0 = the template's size |
| `mark_empty_inputs` | `false`                                | Write `empty_input_text` into empty cells of input (`*`) columns in Word |
| `empty_input_text` | `TBD`                                   | Text for those cells (1–100 characters) |
| `host`, `port`  | `127.0.0.1`, `8000`                        | Where the server listens                 |

Relative paths in the file are resolved against the folder containing it.
Saving settings in the GUI rewrites the file (your own comments in it are lost).
Windows paths: inside double quotes a backslash must be doubled
(`"C:\\Data\\items.xlsx"`); single quotes (`'C:\Data\items.xlsx'`) or forward
slashes (`"C:/Data/items.xlsx"`) also work.
`--host` / `--port` on the command line override the file. The old `DLG_*`
environment variables are no longer used. Settings added by a newer version
are appended to an existing file with their defaults, so you can see and change them.

## Using it

**Settings** (collapsible panel at the top): choose the **Item database**
(.xlsx), **Word template** (.docx), **Table layout** (.xlsx, optional) and
**Output folder** with **Browse…**, set the **ID column**, **Layout sheet** and
**Table font size**, choose whether **Empty input cells** get a text such as
`TBD`, and **Save settings**. They are written to your configuration file and take effect
immediately: the new database is loaded, and selected items that are not in it
are deselected. If the database cannot be loaded at start-up, the panel opens
by itself so the path or ID column can be fixed. Settings and browsing only
work from the computer running the tool.

1. Tick the items that belong on the declared list (click the checkbox or the
   row). The header checkbox selects everything currently shown.
   - **Search** looks in all shown columns (every word must match).
   - **Column filters** work like Excel's: click ▾ in a column header, tick the
     values to show (type in the box to narrow the list; Enter = OK). The list
     only offers values left by the other columns' filters. A filtered column's
     ▾ is highlighted; **Clear filters** removes them all.
   - Search, column filters and **Selected only** all apply together.
     Filters are not saved in the session.
   - **Input columns** (marked ✎) take a value you type per item, e.g. a
     serial number. Their cells are editable for selected rows; clicking the
     cell of an unselected row selects it. Typed values are saved with the
     session, are searchable and filterable, and go into the Word table.
2. Enter a **Name** and, optionally, a **Configuration** (what the items
   belong to, e.g. `Rig 7, build B`) and **Save** to store the selection as a
   session in My sessions. **Save to…** saves it into a folder you choose
   (the shared library, My sessions or any folder) and asks before replacing
   a file there. The page header shows the current session and configuration.
   - **Session library…** lists the sessions of a folder: **My sessions**,
     the **Shared library** (`global_sessions_dir`; **Change folder…** sets
     it) or any **Other folder**. Search by name, configuration or file, sort
     by clicking a column header, and **Load** or **Delete** the selected session.
     Dates are shown as `YYYY-MM-DD HH:MM` (local time), whatever the browser language.
3. Check the **File name** (suggested from the configuration name, else the
   session name; `.docx` is added) and click **Export Word document**. The
   file is written to the output folder (created if needed) and the full path
   is shown. If a file of that name exists you are asked before it is replaced;
   if it is open in Word, close it and export again. Characters not allowed in
   Windows file names (`< > : " / \ | ? *`) are refused. Emptying the field
   and leaving it brings the suggestion back (as does **Reset name**).
   After the export, **Open document** opens it in Word and **Open folder**
   shows it in Explorer (Finder on macOS, the file manager on Linux).

The Excel file is re-read on every page load, so you can edit it and refresh.
The app version is shown in the page header.

## File formats

**Item database (Excel).** First sheet (other than the Word layout sheet),
first row = headers. The columns of the sheet are the columns of the tool:
every column with a header is shown in the browser table and, without a
layout sheet, written to the Word table, in sheet order.

- **ID column:** one column holds a unique ID per item (selections are saved
  by ID). Its header is set by `id_column` in the configuration file (default
  `ID`). Rows without an ID are skipped; duplicate IDs are an error.
- **Omitting a column:** start its header with `#` (e.g. `#Internal note`).
  The tool then ignores it completely: not shown, not searched, not exported.
  Columns without a header are ignored too.
- **User-input column:** end its header with `*` (e.g. `Serial number*`). The
  user types its value per selected item in the tool; the `*` is not shown or
  exported, and the sheet's own cells in that column are ignored. The ID column
  cannot be an input column.
- Two shown columns with the same header are an error (`Serial` and
  `Serial*` count as the same header).
- Dates are shown as `YYYY-MM-DD`; whole numbers without a decimal part.

**Session file (JSON).** One readable file per session in the sessions folder:

```json
{
  "schema_version": 2,
  "name": "Project A",
  "configuration_name": "Rig 7, build B",
  "created": "2026-11-17T18:34:41Z",
  "modified": "2026-11-18T09:12:03Z",
  "source": "items.xlsx",
  "selected_ids": ["DL-001", "DL-003"],
  "user_values": {"DL-001": {"Serial number": "SN-1234"}}
}
```

- `name` identifies the session. The file name is a readable form of it; if
  two names give the same file name (e.g. `a/b` and `a_b`, or `Proj` and `proj`
  on Windows), the second is stored as `a_b (2).json`.
- `configuration_name` describes the configuration the items belong to;
  `user_values` holds the values typed in input columns, for selected items only
  (values of deselected items are dropped on save).
- Files from version 0.1.0 (`schema_version` 1) still load; they are written in
  the new format the next time they are saved. Files from a newer version of
  the tool are refused with a clear message.

**Word template.** Any `.docx`. Put a paragraph containing exactly
`{{DECLARED_LIST}}` where the table should go. The table is inserted there and
inherits that section's page setup, so put the placeholder in a **landscape
section** (with portrait sections before and after, as in the example
template). Without a layout sheet (below), the table has one repeating
header row and exactly the shown (non-`#`) columns of the item database; column
widths follow the length of their content.

The table's font size can be set in Settings → *Table font size* (pt; empty or
0 keeps the template's size and spacing). With a size set, the table cells are
also made compact: no space before/after paragraphs, single line spacing.

With Settings → *Empty input cells* → **Write** ticked, every cell of an input
(`*`) column that was left empty is written as the given text (default `TBD`),
so missing input stands out in the document. Empty database values and blank
layout columns stay empty. The values typed in the tool are not changed.
After every export the status bar lists the empty input cells per column, e.g.
"Empty input cells: Serial number 3 (written as TBD)." (or "left empty" with
the setting off), so missing input is noticed before the document goes out.

`{{CONFIGURATION_NAME}}` anywhere in the template (body text, tables, headers
and footers) is replaced by the configuration name entered in the tool, e.g.
`Configuration: {{CONFIGURATION_NAME}}` as in the example template. Text boxes
are not searched.

**Word table layout (Excel sheet, optional).** Draws a multi-level table header
with merged cells. By default it is a sheet called `Word layout` in the item
database workbook; Settings → *Table layout* / *Layout sheet* can point to
another `.xlsx` and sheet. Example: `data/layout_example.xlsx` (written by
`tools/make_examples.py`).

```
        A        B              C           D           E                F
1   | Item                  | Supplier  | Category  | Traceability                |
2   | No.    | Designation  | (merged ↓)| (merged ↓)| Serial number  | Checked by |
3     ID       Name           Supplier    Category    Serial number*   Checked by*   ← mapping row
```

- The **last used row is the mapping row**. Each cell names a column of the
  data sheet; the selected items are written beneath the header in that order.
  The mapping row itself is **not printed**. The table spans its first to last
  filled cell; a column may be mapped once.
- A **`-`** in the mapping row is a **blank Word column** on purpose: its cells
  stay empty (the header comes from the rows above). It also works at the left
  or right edge of the table.
- The rows **above** it are the Word header: free text; merged cells (across
  and down) become merged cells in Word; Alt+Enter line breaks are kept. With
  no rows above, the mapping names form a one-row header. Merged cells must not
  include the mapping row (merging erases the text of all but the first cell).
- **Only mapped columns** go to Word, in mapping order. Database columns that
  are not mapped are fine: they are still shown in the browser (hide them there
  with `#` as before) but not written to Word.
- A mapping name ending in `*` is a **user-input column**, even if the data
  sheet has it as a normal column. It does not have to exist in the database:
  the tool adds it to the browser table as an input column.
- An **empty mapping cell** inside the table, or a mapping name that is not a
  shown database column (and has no `*`), gives a blank Word column plus a
  **warning** (status bar, ⚠ on the "Word table" label; hover it for details).
  The export still works. Use `-` to silence the warning, or `*` to type the
  column in the tool.
- Mapping names **ignore case and extra spaces** (`part  no` finds `Part No`),
  with a warning so the layout can be corrected. An exact name always wins; a
  name that loosely matches two database columns, or two mapping names that
  land on the same column, are errors.
- **Column widths** in the layout sheet set the relative widths in Word,
  scaled to the page width.
- **Check Word table…** (next to *Export*) shows the header as it will be
  exported: merged cells, column widths in proportion, and a last row naming
  each column's source (input `*` columns in bold, blank columns in red), plus
  any layout warnings. It re-reads the sheet each time, so you can fix the
  layout in Excel, save, and check again without restarting.
- No layout sheet in the database workbook → the automatic table. A separately
  chosen layout file must contain the named sheet.

## Behaviour worth knowing

- After an export, **Open document** opens it in the program registered for
  `.docx` (normally Word) and **Open folder** shows it in its folder (on
  Windows with the file selected). Both only work on the computer running the
  tool.
- Export rows follow **database order**, not the order you clicked.
- Loading a session whose items are no longer in the database skips them and
  lists their IDs in a warning. They are dropped if you save again.

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
  layout.py        Word table layout sheet (multi-level header)
  filebrowser.py   folder listing for the Browse dialog
  docgen.py        Word generation
  opener.py        open the exported document or its folder
  app.py           FastAPI routes + static GUI
  static/          index.html, app.js, logic.js, style.css
tools/make_examples.py   example database, table layout and template
PLAN.md                  current status and plan (decisions in docs/decisions.md)
```

Swapping the item database later means writing another class with a
`name` and `load()` returning an `ItemTable` (see `sources/base.py`) and using it in `app.py`.
