/**
 * Browser GUI for the Declared List Generator.
 *
 * Talks to the FastAPI backend (/api/...). Item text is only ever inserted
 * with textContent, never innerHTML, so spreadsheet content cannot inject markup.
 */
import { errorMessage, filterItems, headerState, partitionIds, setMany } from './logic.js';

const state = {
  columns: [], // shown column names, in database order
  items: [],
  selected: new Set(),
  filter: { search: '', onlySelected: false },
  sessions: [],
  sessionName: '', // name of the session last loaded or saved
  dirty: false, // selection changed since last load/save
};

const $ = (id) => document.getElementById(id);
const el = {
  source: $('source'), status: $('status'), rows: $('rows'), empty: $('empty'),
  search: $('search'), onlySelected: $('only-selected'), headRow: $('head-row'),
  clear: $('clear-btn'), count: $('count'), exportBtn: $('export-btn'),
  selectVisible: $('select-visible'), sessionName: $('session-name'),
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
    throw new Error(errorMessage(body, `Request failed (${res.status})`));
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

/** Items passing the current filter. */
const visible = () => filterItems(state.items, state.filter, state.selected);

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

    tr.append(td, ...state.columns.map((c) => make('td', item.values[c] ?? '')));
    frag.append(tr);
  }
  el.rows.replaceChildren(frag);
  el.empty.hidden = shown.length > 0;
  updateSummary();
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

/** Rebuild the column headers (after the select-all checkbox) from the database columns. */
function renderHeader() {
  const check = el.headRow.firstElementChild;
  el.headRow.replaceChildren(check, ...state.columns.map((c) => {
    const th = make('th', c);
    th.scope = 'col';
    return th;
  }));
}

/** Rebuild the saved-sessions dropdown. */
function renderSessions() {
  const options = state.sessions.map((s) => {
    const when = new Date(s.modified).toLocaleString();
    return new Option(`${s.name} — ${s.count} items — ${when}`, s.name);
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
  const exists = state.sessions.some((s) => s.name.toLowerCase() === name.toLowerCase());
  if (exists && name !== state.sessionName
      && !window.confirm(`A session named "${name}" already exists. Overwrite it?`)) {
    return;
  }
  try {
    const saved = await sendJson('/api/sessions', 'POST', { name, selected_ids: [...state.selected] });
    state.sessionName = saved.name;
    state.dirty = false;
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
    state.dirty = false;
    el.sessionName.value = session.name;
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
    if (state.sessionName === name) state.sessionName = '';
    await refreshSessions();
    showStatus(`Deleted "${name}".`);
  } catch (err) {
    showStatus(err.message, 'error');
  }
}

/** Generate the Word document for the selection and download it. */
async function exportDocument() {
  el.exportBtn.disabled = true;
  try {
    const res = await sendJson('/api/export', 'POST', { selected_ids: [...state.selected] }, true);
    const blob = await res.blob();
    const url = URL.createObjectURL(blob);
    const a = make('a');
    a.href = url;
    a.download = 'declared_list.docx';
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    showStatus(`Exported ${state.selected.size} items.`);
  } catch (err) {
    showStatus(err.message, 'error');
  } finally {
    updateSummary();
  }
}

/** Wire up all event handlers. */
function bindEvents() {
  el.rows.addEventListener('change', (e) => {
    const tr = e.target.closest('tr');
    if (!tr) return;
    changeSelection(setMany(state.selected, [tr.dataset.id], e.target.checked));
    // With "selected only" on, an unchecked row should disappear.
    if (state.filter.onlySelected) renderRows();
    else {
      tr.classList.toggle('selected', e.target.checked);
      updateSummary();
    }
  });
  // Clicking anywhere on a row toggles it; the checkbox's own click is handled by 'change'.
  el.rows.addEventListener('click', (e) => {
    if (e.target.tagName === 'INPUT') return;
    const cb = e.target.closest('tr')?.querySelector('input');
    if (cb) cb.click();
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
  window.addEventListener('beforeunload', (e) => { if (state.dirty) e.preventDefault(); });
}

/** Load items and sessions, then show the page. */
async function init() {
  bindEvents();
  try {
    const data = await api('/api/items');
    state.columns = data.columns.map((c) => c.name);
    state.items = data.items;
    el.source.textContent = `Database: ${data.source}`;
    renderHeader();
    renderRows();
    await refreshSessions();
  } catch (err) {
    showStatus(err.message, 'error');
    updateSummary();
  }
}

init();
