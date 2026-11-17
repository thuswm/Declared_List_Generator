// Unit tests for the GUI's pure helpers. Run with: node --test tests/js
import test from 'node:test';
import assert from 'node:assert/strict';
import {
  categories, errorMessage, filterItems, headerState, partitionIds, setMany,
} from '../../src/dlg/static/logic.js';

const items = [
  { id: 'A1', name: 'Hydraulic Pump', description: 'Main pump', category: 'Hydraulics', extra: { Supplier: 'Acme' } },
  { id: 'B2', name: 'Pressure Sensor', description: 'Transducer', category: 'Sensors', extra: {} },
  { id: 'C3', name: 'Relief Valve', description: 'Pump protection', category: 'Hydraulics' },
];
const none = { search: '', category: '', onlySelected: false };

test('filter: no filter returns everything', () => {
  assert.equal(filterItems(items, none, new Set()).length, 3);
});

test('filter: search is case-insensitive and matches any field incl. extra', () => {
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

test('filter: category and selected-only combine', () => {
  const sel = new Set(['A1', 'B2']);
  const f = { search: '', category: 'Hydraulics', onlySelected: true };
  assert.deepEqual(filterItems(items, f, sel).map((i) => i.id), ['A1']);
});

test('categories: unique, non-empty, sorted', () => {
  assert.deepEqual(categories([...items, { id: 'D', name: 'x', description: '', category: '' }]),
    ['Hydraulics', 'Sensors']);
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
