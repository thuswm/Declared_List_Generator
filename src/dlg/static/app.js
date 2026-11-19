/**
 * Browser GUI for the Declared List Generator.
 *
 * Talks to the FastAPI backend (/api/...). Item text is only ever inserted
 * with textContent, never innerHTML, so spreadsheet content cannot inject markup.
 */
import {
  chooseFilter, currentLabel, errorMessage, filterItems, filterValues, headerState, matchesQuery, partitionIds,
  sessionLabel, setMany, setUserValue, suggestFileName, withUserValues,
} from './logic.js';

/** Label shown for empty cells in a filter list. */
const BLANK_LABEL = '(Blanks)';

const state = {
  columns: [], // shown columns {name, user_input}, in database order
  items: [],
  selected: new Set(),
  // columns: column name -> Set of allowed values (only filtered columns are present).
  filter: { search: '', onlySelected: false, columns: {} },
  // The open filter panel: column, offered values, ticked values, typed text, opening button.
  panel: null,
  sessions: [],
  sessionName: '', // name of the session last loaded or saved
  configurationName: '', // configuration name of that session, as last loaded or saved
  userValues: {}, // values typed in user-input columns: {itemId: {column: value}}
  savedDatabase: '', // item database path + ID column as last saved (detects a change)
  fileNameEdited: false, // the user typed a file name; stop suggesting one
  dirty: false, // selection changed since last load/save
};

const $ = (id) => document.getElementById(id);
const el = {
  source: $('source'), status: $('status'), rows: $('rows'), empty: $('empty'),
  search: $('search'), onlySelected: $('only-selected'), headRow: $('head-row'),
  clearFilters: $('clear-filters-btn'), panel: $('filter-panel'), panelTitle: $('filter-title'),
  panelQuery: $('filter-query'), panelAll: $('filter-all'), panelValues: $('filter-values'),
  panelOk: $('filter-ok'), panelCancel: $('filter-cancel'), panelClear: $('filter-clear'),
  settings: $('settings'), itemsFile: $('items-file'), idColumn: $('id-column'), outputDir: $('output-dir'),
  templateFile: $('template-file'), templateBrowse: $('template-browse'),
  itemsBrowse: $('items-browse'), outputBrowse: $('output-browse'), settingsSave: $('settings-save'),
  configFile: $('config-file'), browse: $('browse-dialog'), browseTitle: $('browse-title'),
  browseRoots: $('browse-roots'), browseUp: $('browse-up'), browsePath: $('browse-path'), browseGo: $('browse-go'),
  browseList: $('browse-list'), browseError: $('browse-error'), browseChoice: $('browse-choice'),
  browseCancel: $('browse-cancel'), browseSelect: $('browse-select'),
  clear: $('clear-btn'), count: $('count'), exportBtn: $('export-btn'),
  selectVisible: $('select-visible'), sessionName: $('session-name'), configName: $('config-name'),
  current: $('current'), fileName: $('file-name'), fileNameReset: $('file-name-reset'), version: $('version'), exitBtn: $('exit-btn'), stopped: $('stopped'),
  sessionList: $('session-list'), save: $('save-btn'), load: $('load-btn'), del: $('delete-btn'),
};

/**
 * Show a message in the status bar.
 * @param {string} text Message text.
 * @param {'ok'|'error'|'warn'} kind Visual style; errors use role=alert.
 */
function showStatus(text, kind = 'ok') {
  el.status.textContent = text;
  el.status.className = `status ${kind}`;
  el.status.setAttribute('role', kind === 'error' ? 'alert' : 'status');
  el.status.hidden = false;
}

/** Hide the status bar. */
function clearStatus() {
  el.status.hidden = true;
}

/**
 * Call the backend and return parsed JSON (or the raw Response when asked).
 * @param {string} path URL path.
 * @param {RequestInit & {raw?: boolean}} [opts] fetch options; `raw` returns the Response.
 * @returns {Promise<*>} Parsed body, or Response when `raw`.
 * @throws {Error} With a readable message for non-2xx responses or network failure.
 */
async function api(path, opts = {}) {
  const { raw, ...init } = opts;
  let res;
  try {
    res = await fetch(path, init);
  } catch {
    throw new Error('Cannot reach the server. Is it still running?');
  }
  if (!res.ok) {
    const body = await res.json().catch(() => null);
    const err = new Error(errorMessage(body, `Request failed (${res.status})`));
    err.status = res.status; // lets callers react to e.g. 409 (file exists)
    throw err;
  }
  if (raw) return res;
  return res.status === 204 ? null : res.json();
}

/**
 * Send JSON to the backend.
 * @param {string} path URL path.
 * @param {string} method HTTP method.
 * @param {*} body JSON-serialisable body.
 * @param {boolean} [raw] Return the Response instead of parsed JSON.
 * @returns {Promise<*>} See {@link api}.
 */
function sendJson(path, method, body, raw = false) {
  return api(path, {
    method, raw,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
}

/**
 * Create an element with optional text and class.
 * @param {string} tag Tag name.
 * @param {string} [text] Text content.
 * @param {string} [cls] Class name.
 * @returns {HTMLElement} The new element.
 */
function make(tag, text = '', cls = '') {
  const node = document.createElement(tag);
  if (text) node.textContent = text;
  if (cls) node.className = cls;
  return node;
}

/** All items with the typed values filled into the user-input columns. */
const mergedItems = () => withUserValues(state.items, state.columns, state.userValues);

/** Items passing the current filter (typed values included). */
const visible = () => filterItems(mergedItems(), state.filter, state.selected);

/** Rebuild the table body for the current filter. */
function renderRows() {
  const shown = visible();
  const frag = document.createDocumentFragment();
  for (const item of shown) {
    const tr = make('tr');
    tr.dataset.id = item.id;
    tr.classList.toggle('selected', state.selected.has(item.id));

    const td = make('td', '', 'check');
    const cb = make('input');
    cb.type = 'checkbox';
    cb.checked = state.selected.has(item.id);
    cb.setAttribute('aria-label', `Select ${item.id}`);
    td.append(cb);

    tr.append(td, ...state.columns.map((c) => (c.user_input ? inputCell(item, c.name) : make('td', item.values[c.name] ?? ''))));
    frag.append(tr);
  }
  el.rows.replaceChildren(frag);
  el.empty.hidden = shown.length > 0;
  updateSummary();
}

/**
 * Build an editable cell for a user-input column. It is read-only while the
 * row is not selected, because only values of selected items are kept.
 * @param {{id:string,values:Object<string,string>}} item The (merged) item.
 * @param {string} column Column name.
 * @returns {HTMLTableCellElement} The cell.
 */
function inputCell(item, column) {
  const td = make('td', '', 'input-cell');
  const input = make('input', '', 'cell-input');
  input.type = 'text';
  input.value = item.values[column] ?? '';
  input.dataset.column = column;
  input.readOnly = !state.selected.has(item.id);
  input.setAttribute('aria-label', `${column} for ${item.id}`);
  td.append(input);
  return td;
}

/** Refresh the count, header checkbox and button states (no row rebuild). */
function updateSummary() {
  const shown = visible();
  const n = state.selected.size;
  el.count.textContent = `${n} selected · showing ${shown.length} of ${state.items.length}`;
  const hs = headerState(shown, state.selected);
  el.selectVisible.checked = hs === 'all';
  el.selectVisible.indeterminate = hs === 'some';
  el.selectVisible.disabled = shown.length === 0;
  el.exportBtn.disabled = n === 0;
  const nFilters = Object.keys(state.filter.columns).length;
  el.clearFilters.textContent = nFilters ? `Clear filters (${nFilters})` : 'Clear filters';
  el.clearFilters.disabled = nFilters === 0;
  el.clear.disabled = n === 0;
  el.del.disabled = !el.sessionList.value;
  el.load.disabled = !el.sessionList.value;
}

/**
 * Replace the selection and mark the session as changed.
 * @param {Set<string>} next The new selection.
 */
function changeSelection(next) {
  state.selected = next;
  state.dirty = true;
}

/** Build the column headers (after the select-all checkbox), each with a filter button. */
function renderHeader() {
  const check = el.headRow.firstElementChild;
  el.headRow.replaceChildren(check, ...state.columns.map(({ name, user_input: userInput }) => {
    const th = make('th');
    th.scope = 'col';
    const inner = make('div', '', 'th-inner');
    const btn = make('button', '▾', 'filter-btn');
    btn.type = 'button';
    btn.dataset.column = name;
    btn.setAttribute('aria-haspopup', 'dialog');
    btn.setAttribute('aria-expanded', 'false');
    const label = make('span', name);
    if (userInput) {
      label.append(make('span', ' ✎', 'input-mark'));
      th.title = 'Typed by you for each selected item';
    }
    inner.append(label, btn);
    th.append(inner);
    return th;
  }));
  updateFilterButtons();
}

/** Highlight the filter buttons of filtered columns. */
function updateFilterButtons() {
  for (const btn of el.headRow.querySelectorAll('.filter-btn')) {
    const active = Object.hasOwn(state.filter.columns, btn.dataset.column);
    btn.classList.toggle('active', active);
    btn.setAttribute('aria-label', `Filter ${btn.dataset.column}${active ? ' (filtered)' : ''}`);
    btn.title = active ? 'Filtered – click to change' : 'Filter this column';
  }
}

/** Apply a change to the column filters and refresh the table. */
function filtersChanged() {
  updateFilterButtons();
  renderRows();
}

/**
 * Open the filter panel for a column, below its header button.
 * @param {HTMLButtonElement} btn The column's filter button.
 */
function openFilter(btn) {
  const column = btn.dataset.column;
  if (state.panel) closeFilter(false);
  const values = filterValues(mergedItems(), state.filter.columns, column);
  const current = state.filter.columns[column];
  state.panel = {
    column, values, button: btn, query: '',
    checked: new Set(current ? values.filter((v) => current.has(v)) : values),
  };
  el.panelTitle.textContent = `Filter: ${column}`;
  el.panelQuery.value = '';
  el.panelClear.disabled = !current;
  renderFilterList();
  el.panel.hidden = false;
  btn.setAttribute('aria-expanded', 'true');
  positionPanel();
  el.panelQuery.focus();
}

/** Place the panel under its button, kept inside the window. */
function positionPanel() {
  const r = state.panel.button.getBoundingClientRect();
  const w = el.panel.offsetWidth;
  el.panel.style.left = `${Math.max(8, Math.min(r.left, window.innerWidth - w - 8))}px`;
  el.panel.style.top = `${Math.min(r.bottom + 4, Math.max(8, window.innerHeight - el.panel.offsetHeight - 8))}px`;
}

/** Values of the open panel that match the typed text. */
const listedValues = () => state.panel.values.filter((v) => matchesQuery(v, state.panel.query, BLANK_LABEL));

/** Rebuild the value checkboxes of the open panel and its select-all/OK state. */
function renderFilterList() {
  const listed = listedValues();
  const frag = document.createDocumentFragment();
  for (const v of listed) {
    const label = make('label', '', 'filter-option');
    const cb = make('input');
    cb.type = 'checkbox';
    cb.checked = state.panel.checked.has(v);
    cb.dataset.value = v;
    label.append(cb, v === '' ? make('span', BLANK_LABEL, 'filter-blank') : make('span', v));
    frag.append(label);
  }
  if (listed.length === 0) frag.append(make('p', 'No matching values.', 'muted filter-option'));
  el.panelValues.replaceChildren(frag);
  updateFilterControls();
}

/** Refresh the select-all box and OK button without rebuilding the list. */
function updateFilterControls() {
  const listed = listedValues();
  const n = listed.filter((v) => state.panel.checked.has(v)).length;
  el.panelAll.checked = listed.length > 0 && n === listed.length;
  el.panelAll.indeterminate = n > 0 && n < listed.length;
  el.panelAll.disabled = listed.length === 0;
  el.panelOk.disabled = n === 0; // an empty filter would hide every row
}

/**
 * Close the filter panel.
 * @param {boolean} [refocus] Return focus to the column's filter button.
 */
function closeFilter(refocus = true) {
  if (!state.panel) return;
  const { button } = state.panel;
  button.setAttribute('aria-expanded', 'false');
  el.panel.hidden = true;
  state.panel = null;
  if (refocus) button.focus();
}

/** Apply the open panel's choice (OK). */
function applyFilter() {
  const { column, values, checked, query } = state.panel;
  const allowed = chooseFilter(values, checked, query);
  if (allowed) state.filter.columns[column] = allowed;
  else delete state.filter.columns[column];
  closeFilter();
  filtersChanged();
}

/** Remove the open panel's column filter. */
function clearColumnFilter() {
  delete state.filter.columns[state.panel.column];
  closeFilter();
  filtersChanged();
}

/** Wire up the column filter buttons and the filter panel. */
function bindFilterEvents() {
  el.headRow.addEventListener('click', (e) => {
    const btn = e.target.closest('.filter-btn');
    if (!btn) return;
    if (state.panel?.button === btn) closeFilter();
    else openFilter(btn);
  });
  el.panelQuery.addEventListener('input', () => {
    state.panel.query = el.panelQuery.value;
    renderFilterList();
  });
  el.panelValues.addEventListener('change', (e) => {
    const v = e.target.dataset.value;
    if (v === undefined) return;
    if (e.target.checked) state.panel.checked.add(v);
    else state.panel.checked.delete(v);
    updateFilterControls();
  });
  el.panelAll.addEventListener('change', () => {
    for (const v of listedValues()) {
      if (el.panelAll.checked) state.panel.checked.add(v);
      else state.panel.checked.delete(v);
    }
    renderFilterList();
  });
  el.panelOk.addEventListener('click', applyFilter);
  el.panelCancel.addEventListener('click', () => closeFilter());
  el.panelClear.addEventListener('click', clearColumnFilter);
  el.panel.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') closeFilter();
    else if (e.key === 'Enter' && e.target === el.panelQuery && !el.panelOk.disabled) applyFilter();
  });
  // Clicking elsewhere cancels, like a menu.
  document.addEventListener('mousedown', (e) => {
    if (state.panel && !el.panel.contains(e.target) && e.target !== state.panel.button) closeFilter(false);
  });
  window.addEventListener('resize', () => closeFilter(false));
  el.clearFilters.addEventListener('click', () => {
    state.filter.columns = {};
    closeFilter(false);
    filtersChanged();
  });
}

/** Show the current session and its configuration in the page header. */
function renderCurrent() {
  const text = currentLabel(state.sessionName, state.configurationName);
  el.current.textContent = text;
  el.current.hidden = !text;
}

/** Rebuild the saved-sessions dropdown. */
function renderSessions() {
  const options = state.sessions.map((s) => {
    const option = new Option(sessionLabel(s), s.name);
    option.title = option.text; // full text on hover when the list is narrow
    return option;
  });
  if (options.length === 0) options.push(new Option('(no saved sessions)', ''));
  el.sessionList.replaceChildren(...options);
  const keep = state.sessions.some((s) => s.name === state.sessionName);
  if (keep) el.sessionList.value = state.sessionName;
  updateSummary();
}

/** Fetch and show the saved-sessions list. */
async function refreshSessions() {
  state.sessions = await api('/api/sessions');
  renderSessions();
}

/**
 * Ask before discarding unsaved selection changes.
 * @returns {boolean} True if it is OK to continue.
 */
function confirmDiscard() {
  return !state.dirty || window.confirm('You have unsaved changes to the current selection. Discard them?');
}

/** Save the current selection under the entered name. */
async function saveSession() {
  const name = el.sessionName.value.trim();
  if (!name) {
    showStatus('Enter a session name first.', 'error');
    el.sessionName.focus();
    return;
  }
  // Names are exact (case-sensitive) keys on the server, so compare exactly here too.
  const exists = state.sessions.some((s) => s.name === name);
  if (exists && name !== state.sessionName
      && !window.confirm(`A session named "${name}" already exists. Overwrite it?`)) {
    return;
  }
  try {
    const saved = await sendJson('/api/sessions', 'POST', {
      name,
      configuration_name: el.configName.value.trim(),
      selected_ids: [...state.selected],
      user_values: state.userValues,
    });
    state.sessionName = saved.name;
    state.configurationName = saved.configuration_name;
    el.configName.value = saved.configuration_name;
    state.dirty = false;
    renderCurrent();
    await refreshSessions();
    showStatus(`Saved "${saved.name}" (${saved.selected_ids.length} items).`);
  } catch (err) {
    showStatus(err.message, 'error');
  }
}

/** Load the session chosen in the dropdown. */
async function loadSession() {
  const name = el.sessionList.value;
  if (!name || !confirmDiscard()) return;
  try {
    const session = await api(`/api/sessions/${encodeURIComponent(name)}`);
    const { known, unknown } = partitionIds(session.selected_ids, state.items);
    state.selected = new Set(known);
    state.sessionName = session.name;
    state.configurationName = session.configuration_name;
    state.userValues = session.user_values;
    state.dirty = false;
    el.sessionName.value = session.name;
    el.configName.value = session.configuration_name;
    useSuggestedName(); // a loaded session brings its own name
    renderCurrent();
    renderRows();
    if (unknown.length) {
      showStatus(`Loaded "${session.name}", but ${unknown.length} item(s) are no longer in the database `
        + `and were skipped: ${unknown.join(', ')}`, 'warn');
    } else {
      showStatus(`Loaded "${session.name}" (${known.length} items).`);
    }
  } catch (err) {
    showStatus(err.message, 'error');
  }
}

/** Delete the session chosen in the dropdown after confirmation. */
async function deleteSession() {
  const name = el.sessionList.value;
  if (!name || !window.confirm(`Delete session "${name}"? This cannot be undone.`)) return;
  try {
    await api(`/api/sessions/${encodeURIComponent(name)}`, { method: 'DELETE' });
    if (state.sessionName === name) {
      state.sessionName = '';
      state.configurationName = '';
      renderCurrent();
    }
    await refreshSessions();
    showStatus(`Deleted "${name}".`);
  } catch (err) {
    showStatus(err.message, 'error');
  }
}

/** Propose a file name from the configuration/session name unless the user typed one. */
function suggestName() {
  if (!state.fileNameEdited) el.fileName.value = suggestFileName(el.configName.value, el.sessionName.value);
  el.fileNameReset.disabled = !state.fileNameEdited;
}

/** Go back to following the configuration/session name (Reset name, session load, emptied field). */
function useSuggestedName() {
  state.fileNameEdited = false;
  suggestName();
}

/**
 * Write the Word document for the selection to the output folder.
 * Asks before replacing an existing file of the same name.
 */
async function exportDocument() {
  el.exportBtn.disabled = true;
  if (!el.fileName.value.trim()) useSuggestedName(); // e.g. Enter pressed in an emptied field
  const body = {
    selected_ids: [...state.selected],
    user_values: state.userValues,
    file_name: el.fileName.value.trim(),
    configuration_name: el.configName.value.trim(),
    overwrite: false,
  };
  try {
    let result;
    try {
      result = await sendJson('/api/export', 'POST', body);
    } catch (err) {
      if (err.status !== 409) throw err;
      if (!window.confirm(`${err.message}\n\nReplace it?`)) {
        showStatus('Export cancelled; the existing file was kept.', 'warn');
        return;
      }
      result = await sendJson('/api/export', 'POST', { ...body, overwrite: true });
    }
    showStatus(`Saved ${result.count} item${result.count === 1 ? '' : 's'} to ${result.path}`);
  } catch (err) {
    showStatus(err.message, 'error');
  } finally {
    updateSummary();
  }
}

/**
 * Put settings into the panel and remember which database they describe.
 * @param {{items_file:string,id_column:string,template_file:string,output_dir:string}} cfg Settings from the server.
 */
function showSettings(cfg) {
  el.itemsFile.value = cfg.items_file;
  el.idColumn.value = cfg.id_column;
  el.templateFile.value = cfg.template_file;
  el.outputDir.value = cfg.output_dir;
  state.savedDatabase = `${cfg.items_file}\n${cfg.id_column}`;
}

/** Fill the settings panel from the server. */
async function loadSettings() {
  try {
    const cfg = await api('/api/settings');
    showSettings(cfg);
    el.configFile.textContent = cfg.config_file ? `Saved in ${cfg.config_file}` : '';
  } catch (err) {
    showStatus(err.message, 'error');
  }
}

/** Save the settings; reload the items if the database changed. */
async function saveSettings() {
  el.settingsSave.disabled = true;
  try {
    const before = state.savedDatabase;
    const cfg = await sendJson('/api/settings', 'PUT', {
      items_file: el.itemsFile.value.trim(),
      id_column: el.idColumn.value.trim(),
      template_file: el.templateFile.value.trim(),
      output_dir: el.outputDir.value.trim(),
    });
    showSettings(cfg);
    if (state.savedDatabase === before && state.items.length) {
      showStatus('Settings saved.');
      return;
    }
    try {
      const dropped = await loadItems();
      if (dropped.length) {
        showStatus(`Settings saved. ${dropped.length} selected item(s) are not in the new database `
          + `and were deselected: ${dropped.join(', ')}`, 'warn');
      } else {
        showStatus(`Settings saved. Loaded ${state.items.length} items from the new database.`);
      }
    } catch (err) {
      state.items = [];
      state.columns = [];
      el.source.textContent = `Database: ${cfg.items_file.split(/[\\/]/).pop()} (not loaded)`;
      renderHeader();
      renderRows();
      showStatus(`Settings saved, but the database cannot be used: ${err.message}`, 'error');
    }
  } catch (err) {
    showStatus(err.message, 'error');
  } finally {
    el.settingsSave.disabled = false;
  }
}

// ---- Browse dialog -------------------------------------------------------

/** State of the open browse dialog. */
const browse = { files: true, ext: '.xlsx', path: '', chosen: '', onChoose: null };

/**
 * Open the browse dialog.
 * @param {{title:string, files:boolean, ext?:string, start:string, onChoose:(path:string)=>void}} opts
 *   `files`: choose a file (true) or a folder (false); `ext`: extension of the files to list
 *   (default `.xlsx`); `start`: folder or file to open at.
 */
function openBrowse({ title, files, ext = '.xlsx', start, onChoose }) {
  Object.assign(browse, { files, ext, chosen: '', onChoose });
  el.browseTitle.textContent = title;
  el.browseSelect.textContent = files ? 'Select file' : 'Select this folder';
  el.browse.showModal();
  navigate(start);
}

/**
 * Show a folder in the dialog.
 * @param {string} path Folder (or file) to show; '' for the home folder.
 */
async function navigate(path) {
  el.browseError.hidden = true;
  try {
    const q = new URLSearchParams({ path, files: String(browse.files), ext: browse.ext });
    const listing = await api(`/api/fs?${q}`);
    browse.path = listing.path;
    // Keep the chosen file when (re)opening at it; otherwise reset the choice.
    if (!browse.files || !listing.files.some((f) => f.path === path)) browse.chosen = '';
    else browse.chosen = path;
    el.browsePath.value = listing.path;
    el.browseUp.disabled = !listing.parent;
    el.browseUp.dataset.path = listing.parent || '';
    el.browseRoots.replaceChildren(new Option('Go to…', ''),
      ...listing.roots.map((r) => new Option(r.name, r.path)));
    renderBrowseList(listing);
  } catch (err) {
    el.browseError.textContent = err.message;
    el.browseError.hidden = false;
  }
  updateBrowseChoice();
}

/**
 * Render the folders and files of a listing.
 * @param {{dirs:Array<{name:string,path:string}>, files:Array<{name:string,path:string}>}} listing
 */
function renderBrowseList(listing) {
  const items = [
    ...listing.dirs.map((d) => ({ ...d, dir: true })),
    ...listing.files.map((f) => ({ ...f, dir: false })),
  ];
  const frag = document.createDocumentFragment();
  for (const entry of items) {
    const li = make('li');
    const btn = make('button', `${entry.dir ? '📁' : '📄'} ${entry.name}`);
    btn.type = 'button';
    btn.dataset.path = entry.path;
    btn.dataset.dir = String(entry.dir);
    btn.classList.toggle('chosen', entry.path === browse.chosen);
    li.append(btn);
    frag.append(li);
  }
  if (items.length === 0) {
    frag.append(make('li', browse.files ? `No folders or ${browse.ext} files here.` : 'No sub-folders here.', 'empty'));
  }
  el.browseList.replaceChildren(frag);
}

/** Show what "Select" will pick and enable it accordingly. */
function updateBrowseChoice() {
  const pick = browse.files ? browse.chosen : browse.path;
  el.browseChoice.textContent = pick || (browse.files ? `Choose a ${browse.ext} file` : '');
  el.browseSelect.disabled = !pick;
}

/** Finish the dialog with the current choice. */
function chooseBrowse() {
  const pick = browse.files ? browse.chosen : browse.path;
  if (!pick) return;
  el.browse.close();
  browse.onChoose(pick);
}

/** Wire up the settings panel and the browse dialog. */
function bindSettingsEvents() {
  el.settingsSave.addEventListener('click', saveSettings);
  el.itemsBrowse.addEventListener('click', () => openBrowse({
    title: 'Choose the item database (.xlsx)', files: true, start: el.itemsFile.value.trim(),
    onChoose: (p) => { el.itemsFile.value = p; el.settingsSave.focus(); },
  }));
  el.templateBrowse.addEventListener('click', () => openBrowse({
    title: 'Choose the Word template (.docx)', files: true, ext: '.docx', start: el.templateFile.value.trim(),
    onChoose: (p) => { el.templateFile.value = p; el.settingsSave.focus(); },
  }));
  el.outputBrowse.addEventListener('click', () => openBrowse({
    title: 'Choose the output folder', files: false, start: el.outputDir.value.trim(),
    onChoose: (p) => { el.outputDir.value = p; el.settingsSave.focus(); },
  }));
  el.browseList.addEventListener('click', (e) => {
    const btn = e.target.closest('button');
    if (!btn) return;
    if (btn.dataset.dir === 'true') {
      navigate(btn.dataset.path);
      return;
    }
    browse.chosen = btn.dataset.path;
    for (const b of el.browseList.querySelectorAll('button')) b.classList.toggle('chosen', b === btn);
    updateBrowseChoice();
  });
  el.browseList.addEventListener('dblclick', (e) => {
    const btn = e.target.closest('button');
    if (btn && btn.dataset.dir === 'false') chooseBrowse();
  });
  el.browseUp.addEventListener('click', () => navigate(el.browseUp.dataset.path));
  el.browseRoots.addEventListener('change', () => { if (el.browseRoots.value) navigate(el.browseRoots.value); });
  el.browseGo.addEventListener('click', () => navigate(el.browsePath.value.trim()));
  el.browsePath.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault(); // Enter would otherwise close the dialog
      navigate(el.browsePath.value.trim());
    }
  });
  el.browseCancel.addEventListener('click', () => el.browse.close());
  el.browseSelect.addEventListener('click', chooseBrowse);
}

/** Stop the server (after a warning about unsaved changes) and show that it has stopped. */
async function exitTool() {
  if (state.dirty && !window.confirm('You have unsaved changes to the current session. Exit anyway?')) return;
  if (!state.dirty && !window.confirm('Stop the Declared List Generator?')) return;
  try {
    await api('/api/shutdown', { method: 'POST' });
  } catch (err) {
    showStatus(err.message, 'error');
    return;
  }
  state.dirty = false; // nothing to protect any more; skip the "leave page?" prompt
  closeFilter(false);
  if (el.browse.open) el.browse.close();
  for (const node of document.body.children) {
    if (node !== el.stopped && node.tagName !== 'SCRIPT') node.hidden = true;
  }
  el.stopped.hidden = false;
  document.title = 'Declared List Generator (stopped)';
  window.close(); // only works for tabs opened by script; harmless otherwise
}

/** Wire up all event handlers. */
function bindEvents() {
  el.exitBtn.addEventListener('click', exitTool);
  bindFilterEvents();
  bindSettingsEvents();
  el.rows.addEventListener('change', (e) => {
    const tr = e.target.closest('tr');
    if (!tr || e.target.type !== 'checkbox') return; // text cells are handled by 'input'
    changeSelection(setMany(state.selected, [tr.dataset.id], e.target.checked));
    // With "selected only" on, an unchecked row should disappear.
    if (state.filter.onlySelected) renderRows();
    else {
      tr.classList.toggle('selected', e.target.checked);
      for (const input of tr.querySelectorAll('.cell-input')) input.readOnly = !e.target.checked;
      updateSummary();
    }
  });
  el.rows.addEventListener('input', (e) => {
    if (!e.target.classList.contains('cell-input')) return;
    const id = e.target.closest('tr').dataset.id;
    state.userValues = setUserValue(state.userValues, id, e.target.dataset.column, e.target.value);
    state.dirty = true;
  });
  // Clicking anywhere on a row toggles it; the checkbox's own click is handled by 'change'.
  // Clicking a value cell of an unselected row selects the row and starts editing instead.
  el.rows.addEventListener('click', (e) => {
    const tr = e.target.closest('tr');
    const cb = tr?.querySelector('input[type="checkbox"]');
    if (!cb || e.target === cb) return;
    if (e.target.classList.contains('cell-input')) {
      if (!cb.checked) {
        cb.click();
        e.target.focus();
      }
      return;
    }
    cb.click();
  });
  el.selectVisible.addEventListener('change', () => {
    const ids = visible().map((i) => i.id);
    changeSelection(setMany(state.selected, ids, el.selectVisible.checked));
    renderRows();
  });
  el.search.addEventListener('input', () => { state.filter.search = el.search.value; renderRows(); });
  el.onlySelected.addEventListener('change', () => {
    state.filter.onlySelected = el.onlySelected.checked;
    renderRows();
  });
  el.clear.addEventListener('click', () => { changeSelection(new Set()); renderRows(); });
  el.sessionList.addEventListener('change', updateSummary);
  el.save.addEventListener('click', saveSession);
  el.load.addEventListener('click', loadSession);
  el.del.addEventListener('click', deleteSession);
  el.exportBtn.addEventListener('click', exportDocument);
  el.sessionName.addEventListener('keydown', (e) => { if (e.key === 'Enter') saveSession(); });
  el.configName.addEventListener('keydown', (e) => { if (e.key === 'Enter') saveSession(); });
  el.configName.addEventListener('input', () => { state.dirty = true; suggestName(); });
  el.sessionName.addEventListener('input', suggestName);
  el.fileName.addEventListener('input', () => {
    // Any typing, including deleting the last character, keeps the user's text. Refilling the
    // suggestion while the field is empty made it snap back to the old name mid-edit (Windows test).
    state.fileNameEdited = true;
    el.fileNameReset.disabled = false;
  });
  // Only a field that is *left* empty goes back to following the configuration/session name.
  el.fileName.addEventListener('blur', () => { if (!el.fileName.value.trim()) useSuggestedName(); });
  el.fileNameReset.addEventListener('click', () => { useSuggestedName(); el.fileName.focus(); });
  el.fileName.addEventListener('keydown', (e) => { if (e.key === 'Enter' && !el.exportBtn.disabled) exportDocument(); });
  window.addEventListener('beforeunload', (e) => { if (state.dirty) e.preventDefault(); });
}

/**
 * Fetch the item database and show it. Selected items that no longer exist
 * are dropped (with a warning), and filters on vanished columns are removed.
 * @returns {Promise<string[]>} IDs dropped from the selection.
 */
async function loadItems() {
  const data = await api('/api/items');
  state.columns = data.columns;
  state.items = data.items;
  const names = new Set(data.columns.map((c) => c.name));
  for (const column of Object.keys(state.filter.columns)) {
    if (!names.has(column)) delete state.filter.columns[column];
  }
  const { known, unknown } = partitionIds([...state.selected], state.items);
  if (unknown.length) changeSelection(new Set(known));
  el.source.textContent = `Database: ${data.source}`;
  renderHeader();
  renderRows();
  return unknown;
}

/** Show the app version in the header, so the user can see which version the browser runs. */
async function showVersion() {
  try {
    const info = await api('/api/info');
    el.version.textContent = `v${info.version}`;
  } catch {
    // Cosmetic only; errors from the real requests are reported by them.
  }
}

/** Load items and sessions, then show the page. */
async function init() {
  bindEvents();
  suggestName();
  showVersion(); // not awaited: the header text is cosmetic
  // Settings first and independently: if the database path is wrong, they are how to fix it.
  await loadSettings();
  try {
    await loadItems();
  } catch (err) {
    showStatus(`${err.message} – check the item database in Settings.`, 'error');
    el.settings.open = true;
    updateSummary();
  }
  try {
    await refreshSessions();
  } catch (err) {
    showStatus(err.message, 'error');
  }
}

init();
