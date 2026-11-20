// Unit tests for the GUI's pure helpers. Run with: node --test tests/js
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  chooseFilter, currentLabel, emptyInputsText, errorMessage, filterItems, filterValues, headerState, matchesQuery, partitionIds,
  filterSessions, layoutWarningText, sessionLabel, setMany, setUserValue, sortSessions, suggestFileName, withUserValues,
} from '../../src/dlg/static/logic.js';

const row = (id, name, description, category, supplier = '') => (
  { id, values: { ID: id, Name: name, Description: description, Category: category, Supplier: supplier } });
const items = [
  row('A1', 'Hydraulic Pump', 'Main pump', 'Hydraulics', 'Acme'),
  row('B2', 'Pressure Sensor', 'Transducer', 'Sensors'),
  row('C3', 'Relief Valve', 'Pump protection', 'Hydraulics'),
];
const none = { search: '', onlySelected: false };

test('filter: no filter returns everything', () => {
  assert.equal(filterItems(items, none, new Set()).length, 3);
});

test('filter: search is case-insensitive and matches any column', () => {
  const ids = (f) => filterItems(items, { ...none, search: f }, new Set()).map((i) => i.id);
  assert.deepEqual(ids('HYDRAULIC'), ['A1', 'C3']);
  assert.deepEqual(ids('acme'), ['A1']);
  assert.deepEqual(ids('b2'), ['B2']);
});

test('filter: every search term must match (AND)', () => {
  const f = { ...none, search: 'pump  hydraulics' };
  assert.deepEqual(filterItems(items, f, new Set()).map((i) => i.id), ['A1', 'C3']);
  assert.deepEqual(filterItems(items, { ...none, search: 'pump sensors' }, new Set()), []);
});

test('filter: search and selected-only combine', () => {
  const sel = new Set(['A1', 'B2']);
  const f = { search: 'hydraulics', onlySelected: true };
  assert.deepEqual(filterItems(items, f, sel).map((i) => i.id), ['A1']);
});

test('filter: works with arbitrary column names', () => {
  const custom = [{ id: 'P1', values: { 'Part No': 'P1', Mass: '12 kg' } }, { id: 'P2', values: { 'Part No': 'P2' } }];
  assert.deepEqual(filterItems(custom, { ...none, search: '12 KG' }, new Set()).map((i) => i.id), ['P1']);
});

test('headerState', () => {
  assert.equal(headerState([], new Set(['A1'])), 'none');
  assert.equal(headerState(items, new Set()), 'none');
  assert.equal(headerState(items, new Set(['A1'])), 'some');
  assert.equal(headerState(items, new Set(['A1', 'B2', 'C3', 'Z'])), 'all');
});

test('setMany does not mutate its input', () => {
  const before = new Set(['A1']);
  const added = setMany(before, ['B2'], true);
  assert.deepEqual([...before], ['A1']);
  assert.deepEqual([...added].sort(), ['A1', 'B2']);
  assert.deepEqual([...setMany(added, ['A1', 'Q'], false)], ['B2']);
});

test('partitionIds keeps order and separates unknown IDs', () => {
  assert.deepEqual(partitionIds(['C3', 'X', 'A1', 'Y'], items),
    { known: ['C3', 'A1'], unknown: ['X', 'Y'] });
});

test('errorMessage handles string, list and missing detail', () => {
  assert.equal(errorMessage({ detail: 'Nope' }, 'fb'), 'Nope');
  assert.equal(errorMessage({ detail: [{ loc: ['body', 'name'], msg: 'too short' }] }, 'fb'), 'name: too short');
  assert.equal(errorMessage(null, 'fb'), 'fb');
});

const ids = (list) => list.map((i) => i.id);

test('column filters: only allowed values pass; blank cells count as ""', () => {
  const columns = { Category: new Set(['Hydraulics']) };
  assert.deepEqual(ids(filterItems(items, { ...none, columns }, new Set())), ['A1', 'C3']);
  const blanks = { Supplier: new Set(['']) };
  assert.deepEqual(ids(filterItems(items, { ...none, columns: blanks }, new Set())), ['B2', 'C3']);
  const missing = [{ id: 'X', values: {} }];
  assert.deepEqual(ids(filterItems(missing, { ...none, columns: blanks }, new Set())), ['X']);
});

test('column filters combine with each other, search and selected-only (AND)', () => {
  const columns = { Category: new Set(['Hydraulics', 'Sensors']), Supplier: new Set(['']) };
  assert.deepEqual(ids(filterItems(items, { ...none, columns }, new Set())), ['B2', 'C3']);
  assert.deepEqual(ids(filterItems(items, { ...none, columns, search: 'valve' }, new Set())), ['C3']);
  assert.deepEqual(ids(filterItems(items, { ...none, columns, onlySelected: true }, new Set(['B2']))), ['B2']);
});

test('filterValues: distinct, natural order, blanks last', () => {
  const list = [
    { id: '1', values: { Pos: 'A10' } }, { id: '2', values: { Pos: 'a2' } },
    { id: '3', values: { Pos: '' } }, { id: '4', values: { Pos: 'A10' } }, { id: '5', values: {} },
  ];
  assert.deepEqual(filterValues(list, {}, 'Pos'), ['a2', 'A10', '']);
});

test('filterValues: cascades from the other columns, ignoring its own filter', () => {
  const columns = { Category: new Set(['Sensors']), Supplier: new Set(['Acme']) };
  // Supplier list is narrowed by the Category filter only, not by its own filter.
  assert.deepEqual(filterValues(items, columns, 'Supplier'), ['']);
  assert.deepEqual(filterValues(items, { Category: new Set(['Sensors']) }, 'Category'),
    ['Hydraulics', 'Sensors']);
});

test('matchesQuery: case-insensitive substring; blanks match their label', () => {
  assert.equal(matchesQuery('Hydraulics', 'DRAU'), true);
  assert.equal(matchesQuery('Hydraulics', '  '), true);
  assert.equal(matchesQuery('Sensors', 'hyd'), false);
  assert.equal(matchesQuery('', 'blank'), true);
  assert.equal(matchesQuery('', 'x'), false);
});

test('chooseFilter: all ticked means no filter; typed text keeps only listed ticked values', () => {
  const values = ['Electrical', 'Hydraulics', 'Sensors', ''];
  assert.equal(chooseFilter(values, new Set(values), ''), null);
  assert.deepEqual([...chooseFilter(values, new Set(['Sensors', '']), '')], ['Sensors', '']);
  // Excel behaviour: with search text, unlisted values are dropped even if ticked.
  assert.deepEqual([...chooseFilter(values, new Set(values), 'sens')], ['Sensors']);
  assert.deepEqual([...chooseFilter(values, new Set(), '')], []);
});

test('sessionLabel shows the configuration name when set', () => {
  const fmt = () => 'DATE';
  assert.equal(sessionLabel({ name: 'P', configuration_name: 'Rig 7', count: 3, modified: 'x' }, fmt),
    'P — Rig 7 — 3 items — DATE');
  assert.equal(sessionLabel({ name: 'P', configuration_name: '', count: 1, modified: 'x' }, fmt),
    'P — 1 item — DATE');
  assert.equal(sessionLabel({ name: 'P', count: 0, modified: 'x' }, fmt), 'P — 0 items — DATE');
});

test('currentLabel', () => {
  assert.equal(currentLabel('', 'Rig'), '');
  assert.equal(currentLabel('P', ''), 'Session: P');
  assert.equal(currentLabel('P', 'Rig 7'), 'Session: P · Configuration: Rig 7');
});

test('withUserValues fills only input columns and does not mutate', () => {
  const cols = [{ name: 'ID' }, { name: 'SN', user_input: true }];
  const list = [{ id: '1', values: { ID: '1', SN: '' } }, { id: '2', values: { ID: '2', SN: '' } }];
  const merged = withUserValues(list, cols, { 1: { SN: 'S-1', ID: 'x' } });
  assert.deepEqual(merged.map((i) => i.values), [{ ID: '1', SN: 'S-1' }, { ID: '2', SN: '' }]);
  assert.equal(list[0].values.SN, '');
  assert.equal(withUserValues(list, [{ name: 'ID' }], {}), list); // no input columns: same array
});

test('typed values are searchable and filterable', () => {
  const cols = [{ name: 'ID' }, { name: 'SN', user_input: true }];
  const list = [{ id: '1', values: { ID: '1', SN: '' } }, { id: '2', values: { ID: '2', SN: '' } }];
  const merged = withUserValues(list, cols, { 2: { SN: 'abc-77' } });
  assert.deepEqual(ids(filterItems(merged, { ...none, search: 'ABC' }, new Set())), ['2']);
  assert.deepEqual(filterValues(merged, {}, 'SN'), ['abc-77', '']);
});

test('setUserValue adds, changes and removes values immutably', () => {
  const a = setUserValue({}, 'X', 'SN', '1');
  assert.deepEqual(a, { X: { SN: '1' } });
  const b = setUserValue(a, 'X', 'Lot', 'L');
  assert.deepEqual(b, { X: { SN: '1', Lot: 'L' } });
  assert.deepEqual(a, { X: { SN: '1' } });
  assert.deepEqual(setUserValue(setUserValue(b, 'X', 'SN', ''), 'X', 'Lot', ''), {});
});

test('suggestFileName: configuration first, then session, then default; Windows-safe', () => {
  assert.equal(suggestFileName('Rig 7, build B', 'Proj'), 'Rig 7, build B');
  assert.equal(suggestFileName('  ', 'Proj A'), 'Proj A');
  assert.equal(suggestFileName('', ''), 'declared_list');
  assert.equal(suggestFileName('a/b:c*d?', ''), 'a_b_c_d_');
  assert.equal(suggestFileName('Rig 7. ', ''), 'Rig 7');
  assert.equal(suggestFileName('...', ''), 'declared_list');
  assert.equal(suggestFileName('x'.repeat(150), '').length, 100);
});

const LIB = [
  { name: 'Rig 10', configuration_name: 'Build B', file: 'Rig 10.json', source: 'items.xlsx', count: 2, modified: '2026-11-18T10:00:00Z' },
  { name: 'rig 2', configuration_name: '', file: 'rig 2.json', source: 'other.xlsx', count: 12, modified: '2026-11-19T09:00:00Z' },
  { name: 'Alpha', configuration_name: 'Build A', file: 'Alpha.json', source: 'items.xlsx', count: 2, modified: '2026-11-17T08:00:00Z' },
];

test('filterSessions: terms AND over name, configuration, file and database', () => {
  assert.deepEqual(filterSessions(LIB, '').map((s) => s.name), ['Rig 10', 'rig 2', 'Alpha']);
  assert.deepEqual(filterSessions(LIB, 'RIG').map((s) => s.name), ['Rig 10', 'rig 2']);
  assert.deepEqual(filterSessions(LIB, 'rig build').map((s) => s.name), ['Rig 10']);
  assert.deepEqual(filterSessions(LIB, 'other.xlsx').map((s) => s.name), ['rig 2']);
  assert.deepEqual(filterSessions(LIB, 'alpha.json').map((s) => s.name), ['Alpha']);
});

test('sortSessions: natural text order, numbers, dates, direction, ties newest first', () => {
  assert.deepEqual(sortSessions(LIB, 'name', 'asc').map((s) => s.name), ['Alpha', 'rig 2', 'Rig 10']);
  assert.deepEqual(sortSessions(LIB, 'name', 'desc').map((s) => s.name), ['Rig 10', 'rig 2', 'Alpha']);
  assert.deepEqual(sortSessions(LIB, 'modified', 'desc').map((s) => s.name), ['rig 2', 'Rig 10', 'Alpha']);
  assert.deepEqual(sortSessions(LIB, 'count', 'asc').map((s) => s.name), ['Rig 10', 'Alpha', 'rig 2']);
  // Blank configuration sorts last, as in the column filter lists.
  assert.deepEqual(sortSessions(LIB, 'configuration_name', 'asc').map((s) => s.name), ['Alpha', 'Rig 10', 'rig 2']);
  assert.equal(LIB[0].name, 'Rig 10'); // input not modified
});

test('layoutWarningText: empty without warnings, counted and joined otherwise', () => {
  assert.equal(layoutWarningText([]), '');
  assert.equal(layoutWarningText(undefined), '');
  assert.equal(layoutWarningText(['A.']), 'Word table: 1 warning – A.');
  assert.equal(layoutWarningText(['A.', 'B.']), 'Word table: 2 warnings – A. B.');
});

test('emptyInputsText: nothing empty, left empty, written as text', () => {
  assert.equal(emptyInputsText({}, 'TBD'), '');
  assert.equal(emptyInputsText(undefined, ''), '');
  assert.equal(emptyInputsText({ Lot: 2, Note: 1 }, ''), 'Empty input cells: Lot 2, Note 1 (left empty).');
  assert.equal(emptyInputsText({ Lot: 2 }, 'TBD'), 'Empty input cells: Lot 2 (written as TBD).');
});
