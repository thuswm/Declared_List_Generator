/**
 * Pure, DOM-free helpers for the Declared List Generator GUI.
 *
 * Kept separate from app.js so they can be unit-tested with Node's built-in
 * test runner (see tests/js/logic.test.js).
 */

/**
 * Build the lower-cased text an item is searched against.
 * @param {{id:string,name:string,description:string,category:string,extra?:Object<string,string>}} item
 * @returns {string} Space-joined searchable text.
 */
function searchText(item) {
  const extra = Object.values(item.extra || {});
  return [item.id, item.name, item.description, item.category, ...extra].join(' ').toLowerCase();
}

/**
 * Filter items by search text, category and "selected only".
 * Every whitespace-separated search term must appear somewhere in the item.
 * @param {Array} items All items.
 * @param {{search:string,category:string,onlySelected:boolean}} filter Active filter.
 * @param {Set<string>} selected IDs of selected items.
 * @returns {Array} Items that pass the filter, in original order.
 */
export function filterItems(items, filter, selected) {
  const terms = filter.search.toLowerCase().split(/\s+/).filter(Boolean);
  return items.filter((item) => {
    if (filter.category && item.category !== filter.category) return false;
    if (filter.onlySelected && !selected.has(item.id)) return false;
    if (terms.length === 0) return true;
    const text = searchText(item);
    return terms.every((t) => text.includes(t));
  });
}

/**
 * List the distinct, non-empty categories.
 * @param {Array} items All items.
 * @returns {string[]} Sorted category names.
 */
export function categories(items) {
  return [...new Set(items.map((i) => i.category).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b),
  );
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
