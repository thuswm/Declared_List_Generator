/**
 * Pure, DOM-free helpers for the Declared List Generator GUI.
 *
 * Kept separate from app.js so they can be unit-tested with Node's built-in
 * test runner (see tests/js/logic.test.js).
 */

/**
 * Build the lower-cased text an item is searched against (all shown columns).
 * @param {{id:string,values:Object<string,string>}} item
 * @returns {string} Space-joined searchable text.
 */
function searchText(item) {
  return [item.id, ...Object.values(item.values || {})].join(' ').toLowerCase();
}

/**
 * Check an item against the column filters.
 * @param {{values:Object<string,string>}} item The item.
 * @param {Object<string,Set<string>>} columnFilters Allowed values per filtered column.
 * @param {string} [except] Column to ignore (used to build that column's value list).
 * @returns {boolean} True if the item's value is allowed in every filtered column.
 */
function passesColumnFilters(item, columnFilters, except) {
  for (const [column, allowed] of Object.entries(columnFilters || {})) {
    if (column === except) continue;
    if (!allowed.has(item.values[column] ?? '')) return false;
  }
  return true;
}

/**
 * Filter items by column filters, search text and "selected only" (all must pass).
 * Every whitespace-separated search term must appear somewhere in the item.
 * @param {Array} items All items.
 * @param {{search:string,onlySelected:boolean,columns?:Object<string,Set<string>>}} filter
 *   Active filter; `columns` maps a column name to its allowed values (absent = unfiltered).
 * @param {Set<string>} selected IDs of selected items.
 * @returns {Array} Items that pass the filter, in original order.
 */
export function filterItems(items, filter, selected) {
  const terms = filter.search.toLowerCase().split(/\s+/).filter(Boolean);
  return items.filter((item) => {
    if (filter.onlySelected && !selected.has(item.id)) return false;
    if (!passesColumnFilters(item, filter.columns)) return false;
    if (terms.length === 0) return true;
    const text = searchText(item);
    return terms.every((t) => text.includes(t));
  });
}

/**
 * Sort cell values for a filter list: natural order ("A2" before "A10"), blanks last.
 * @param {string} a First value.
 * @param {string} b Second value.
 * @returns {number} Comparison result.
 */
function compareValues(a, b) {
  if (a === '' || b === '') return (a === '') - (b === '');
  return a.localeCompare(b, undefined, { numeric: true, sensitivity: 'base' });
}

/**
 * List the values to offer in a column's filter, like Excel: only values of
 * items that pass the filters on the *other* columns.
 * @param {Array} items All items.
 * @param {Object<string,Set<string>>} columnFilters Active column filters.
 * @param {string} column The column whose filter is being opened.
 * @returns {string[]} Distinct values, sorted; '' (blank) last if present.
 */
export function filterValues(items, columnFilters, column) {
  const values = new Set();
  for (const item of items) {
    if (passesColumnFilters(item, columnFilters, column)) values.add(item.values[column] ?? '');
  }
  return [...values].sort(compareValues);
}

/**
 * Check whether a value matches the text typed in a filter panel.
 * @param {string} value Cell value ('' = blank).
 * @param {string} query Typed text.
 * @param {string} blankLabel Label shown for blank values, so typing it finds them.
 * @returns {boolean} True if the value should be listed.
 */
export function matchesQuery(value, query, blankLabel = '(Blanks)') {
  const q = query.trim().toLowerCase();
  if (!q) return true;
  return (value === '' ? blankLabel : value).toLowerCase().includes(q);
}

/**
 * Work out a column's new filter when OK is pressed in its filter panel.
 * As in Excel, only values that are both listed (match the typed text) and
 * checked are kept.
 * @param {string[]} values All values offered for the column.
 * @param {Set<string>} checked Values whose box is ticked.
 * @param {string} query Text typed in the panel's search box.
 * @returns {Set<string>|null} Allowed values, or null when nothing is filtered out.
 */
export function chooseFilter(values, checked, query) {
  const allowed = new Set(values.filter((v) => checked.has(v) && matchesQuery(v, query)));
  return allowed.size === values.length ? null : allowed;
}

/**
 * Work out the state of the "select all visible" header checkbox.
 * @param {Array} visible Currently visible items.
 * @param {Set<string>} selected IDs of selected items.
 * @returns {'none'|'some'|'all'} How many of the visible items are selected.
 */
export function headerState(visible, selected) {
  if (visible.length === 0) return 'none';
  const n = visible.filter((i) => selected.has(i.id)).length;
  if (n === 0) return 'none';
  return n === visible.length ? 'all' : 'some';
}

/**
 * Return a new selection with the given IDs added or removed.
 * @param {Set<string>} selected Current selection (not modified).
 * @param {string[]} ids IDs to change.
 * @param {boolean} on True to select, false to deselect.
 * @returns {Set<string>} The updated selection.
 */
export function setMany(selected, ids, on) {
  const next = new Set(selected);
  for (const id of ids) {
    if (on) next.add(id);
    else next.delete(id);
  }
  return next;
}

/**
 * Split session IDs into those still in the database and those that are not.
 * @param {string[]} ids IDs stored in a session.
 * @param {Array} items All current items.
 * @returns {{known:string[],unknown:string[]}} The partition, order preserved.
 */
export function partitionIds(ids, items) {
  const present = new Set(items.map((i) => i.id));
  return {
    known: ids.filter((id) => present.has(id)),
    unknown: ids.filter((id) => !present.has(id)),
  };
}

/**
 * Turn an error response body into a readable message.
 * FastAPI sends `detail` as a string, or as a list for validation errors.
 * @param {*} body Parsed JSON body (may be null).
 * @param {string} fallback Message to use when nothing better is available.
 * @returns {string} Human-readable message.
 */
export function errorMessage(body, fallback) {
  const detail = body && body.detail;
  if (typeof detail === 'string') return detail;
  if (Array.isArray(detail)) {
    return detail.map((d) => `${(d.loc || []).slice(1).join('.')}: ${d.msg}`).join('; ');
  }
  return fallback;
}

/**
 * Text for a saved session in the load list.
 * @param {{name:string,configuration_name?:string,count:number,modified:string}} s Session summary.
 * @param {(iso:string) => string} [formatDate] Date formatter (injectable for tests).
 * @returns {string} e.g. "Project A — Rig 7 — 3 items — 18/11/2026, 09:12".
 */
export function sessionLabel(s, formatDate = (iso) => new Date(iso).toLocaleString()) {
  const parts = [s.name];
  if (s.configuration_name) parts.push(s.configuration_name);
  parts.push(`${s.count} item${s.count === 1 ? '' : 's'}`, formatDate(s.modified));
  return parts.join(' — ');
}

/**
 * Text for the page header describing the current session.
 * @param {string} sessionName Name of the loaded/saved session ('' = none).
 * @param {string} configurationName Its configuration name ('' = none).
 * @returns {string} Header text, or '' when no session is current.
 */
export function currentLabel(sessionName, configurationName) {
  if (!sessionName) return '';
  return configurationName
    ? `Session: ${sessionName} · Configuration: ${configurationName}`
    : `Session: ${sessionName}`;
}

/**
 * Fill the user's typed values into the user-input columns, so search,
 * filters and the table all see them. Inputs are not modified.
 * @param {Array} items All items (input columns hold '').
 * @param {Array<{name:string,user_input?:boolean}>} columns Shown columns.
 * @param {Object<string,Object<string,string>>} userValues Typed values per item ID and column.
 * @returns {Array} Items with merged values (the same objects if there are no input columns).
 */
export function withUserValues(items, columns, userValues) {
  const inputs = columns.filter((c) => c.user_input).map((c) => c.name);
  if (inputs.length === 0) return items;
  return items.map((item) => {
    const typed = userValues[item.id] || {};
    const values = { ...item.values };
    for (const c of inputs) values[c] = typed[c] ?? '';
    return { ...item, values };
  });
}

/**
 * Return new user values with one value changed (empty values are removed).
 * @param {Object<string,Object<string,string>>} userValues Current values (not modified).
 * @param {string} id Item ID.
 * @param {string} column Column name.
 * @param {string} value Typed value.
 * @returns {Object<string,Object<string,string>>} Updated values.
 */
export function setUserValue(userValues, id, column, value) {
  const row = { ...(userValues[id] || {}) };
  if (value === '') delete row[column];
  else row[column] = value;
  const next = { ...userValues };
  if (Object.keys(row).length) next[id] = row;
  else delete next[id];
  return next;
}

/** Characters Windows does not allow in file names (the server rejects them too). */
const BAD_FILENAME_CHARS = /[<>:"/\\|?*\u0000-\u001f]/g;

/**
 * Suggest an output file name (without .docx) from the configuration or session name.
 * @param {string} configurationName Configuration name ('' if none).
 * @param {string} sessionName Session name ('' if none).
 * @returns {string} A file name that is valid on Windows and Linux.
 */
export function suggestFileName(configurationName, sessionName) {
  const base = (configurationName.trim() || sessionName.trim() || 'declared_list');
  const cleaned = base.replace(BAD_FILENAME_CHARS, '_').slice(0, 100).replace(/[ .]+$/, '');
  return cleaned || 'declared_list';
}

/**
 * Filter the session library by search text. Every whitespace-separated term
 * must appear in the name, configuration, file name or database (case-insensitive).
 * @param {Array<{name:string,configuration_name:string,file:string,source:string}>} sessions Library entries.
 * @param {string} query Typed search text.
 * @returns {Array} Matching entries, in the given order.
 */
export function filterSessions(sessions, query) {
  const terms = query.toLowerCase().split(/\s+/).filter(Boolean);
  if (terms.length === 0) return sessions;
  return sessions.filter((s) => {
    const text = [s.name, s.configuration_name, s.file, s.source].join(' ').toLowerCase();
    return terms.every((t) => text.includes(t));
  });
}

/**
 * Sort session library entries by one column. Text sorts naturally and
 * case-insensitively ("Rig 2" before "Rig 10"); ties keep the newest first.
 * @param {Array} sessions Library entries (not modified).
 * @param {'name'|'configuration_name'|'count'|'modified'|'file'} key Column to sort by.
 * @param {'asc'|'desc'} dir Direction.
 * @returns {Array} A sorted copy.
 */
export function sortSessions(sessions, key, dir) {
  const sign = dir === 'desc' ? -1 : 1;
  const cmp = (a, b) => {
    if (key === 'count') return a.count - b.count;
    // ISO timestamps from the server sort correctly as text.
    if (key === 'modified') return a.modified < b.modified ? -1 : a.modified > b.modified ? 1 : 0;
    return compareValues(a[key] ?? '', b[key] ?? '');
  };
  return [...sessions].sort((a, b) => sign * cmp(a, b) || (a.modified < b.modified ? 1 : a.modified > b.modified ? -1 : 0));
}
