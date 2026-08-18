import assert from 'node:assert/strict';
import test from 'node:test';

import {
  encodeKey,
  equalValues,
  matchesFilter,
  normalizeValue,
  printValue
} from '../src/normalize.js';

test('normalizes blanks and preserves empty strings unless configured otherwise', () => {
  assert.deepEqual(normalizeValue(null), ['blank', null]);
  assert.deepEqual(normalizeValue(undefined), ['blank', null]);
  assert.deepEqual(normalizeValue(''), ['string', '']);
  assert.deepEqual(normalizeValue('', { emptyEqualsNull: true }), ['blank', null]);
  assert.equal(equalValues(normalizeValue(null), normalizeValue('')), false);
});

test('keeps numeric-looking strings distinct from numbers', () => {
  assert.deepEqual(normalizeValue('001'), ['string', '001']);
  assert.deepEqual(normalizeValue(1), ['number', 1]);
  assert.equal(equalValues(normalizeValue('001'), normalizeValue(1)), false);
});

test('applies explicit string trimming and case folding only', () => {
  assert.deepEqual(normalizeValue(' AＢC ', { trim: true, caseSensitive: false }), ['string', 'aｂc']);
  assert.deepEqual(normalizeValue(' A ', { caseSensitive: false }), ['string', ' a ']);
});

test('normalizes supported scalar and ExcelJS cell values', () => {
  assert.deepEqual(normalizeValue(new Date('2026-02-03T04:05:06-05:00')), ['date', '2026-02-03T09:05:06.000Z']);
  assert.deepEqual(normalizeValue(true), ['boolean', true]);
  assert.deepEqual(normalizeValue({ error: '#N/A' }), ['error', '#N/A']);
  assert.deepEqual(normalizeValue({ richText: [{ text: 'Hello' }, { text: ' world' }] }), ['string', 'Hello world']);
  assert.deepEqual(normalizeValue({ foo: 'bar' }), ['string', '[object Object]']);
});

test('normalizes formulas by all supported formula modes', () => {
  const value = { formula: 'A1+1', result: 2 };
  assert.deepEqual(normalizeValue(value, { formulaMode: 'formula' }), ['formula', 'A1+1']);
  assert.deepEqual(normalizeValue(value, { formulaMode: 'cached-result' }), ['number', 2]);
  assert.deepEqual(normalizeValue(value, { formulaMode: 'formula-and-cached-result' }), ['formula', ['A1+1', ['number', 2]]]);
  assert.deepEqual(normalizeValue(value), ['formula', ['A1+1', ['number', 2]]]);
  const shared = { sharedFormula: 'A1', result: 2 };
  assert.deepEqual(normalizeValue(shared, { formulaMode: 'formula' }), ['formula', 'shared:A1']);
  assert.deepEqual(normalizeValue(shared, { formulaMode: 'cached-result' }), ['number', 2]);
  assert.deepEqual(normalizeValue(shared), ['formula', ['shared:A1', ['number', 2]]]);
});

test('normalizes negative zero deterministically', () => {
  const negativeZero = normalizeValue(-0);
  const zero = normalizeValue(0);
  assert.deepEqual(negativeZero, ['number', 0]);
  assert.equal(equalValues(negativeZero, zero), true);
  assert.equal(encodeKey([negativeZero]), encodeKey([zero]));
});

test('rejects non-finite numbers', () => {
  assert.throws(() => normalizeValue(NaN), TypeError);
  assert.throws(() => normalizeValue(Infinity), TypeError);
  assert.throws(() => normalizeValue(-Infinity), TypeError);
});

test('uses numeric tolerance only between numeric values', () => {
  assert.equal(equalValues(['number', 1], ['number', 1.5], { numericTolerance: 0.5 }), true);
  assert.equal(equalValues(['number', 1], ['number', 1.5], { numericTolerance: 0.499 }), false);
  assert.equal(equalValues(['string', 'a'], ['string', 'A'], { numericTolerance: 10 }), false);
});

test('encodes composite typed keys without delimiter collisions', () => {
  assert.notEqual(
    encodeKey([['string', 'a|b'], ['string', 'c']]),
    encodeKey([['string', 'a'], ['string', 'b|c']])
  );
});

test('matches equality, membership, and null filters using typed equality', () => {
  assert.equal(matchesFilter(normalizeValue('001'), { operator: 'eq', value: 1 }), false);
  assert.equal(matchesFilter(normalizeValue(1), { operator: 'ne', value: 1 }), false);
  assert.equal(matchesFilter(normalizeValue(1), { operator: 'ne', value: '1' }), true);
  assert.equal(matchesFilter(normalizeValue('active'), { operator: 'in', value: ['pending', 'active'] }), true);
  assert.equal(matchesFilter(normalizeValue('active'), { operator: 'notIn', values: ['pending', 'closed'] }), true);
  assert.equal(matchesFilter(normalizeValue(null), { operator: 'isNull' }), true);
  assert.equal(matchesFilter(normalizeValue(''), { operator: 'isNotNull' }), true);
  assert.equal(matchesFilter(normalizeValue('', { emptyEqualsNull: true }), { operator: 'isNull' }, { emptyEqualsNull: true }), true);
});

test('matches ordered filters only for comparable values of the same type', () => {
  assert.equal(matchesFilter(normalizeValue(2), { operator: 'gt', value: 1 }), true);
  assert.equal(matchesFilter(normalizeValue(2), { operator: 'gte', value: 2 }), true);
  assert.equal(matchesFilter(normalizeValue(2), { operator: 'lt', value: 3 }), true);
  assert.equal(matchesFilter(normalizeValue(2), { operator: 'lte', value: 2 }), true);
  assert.equal(matchesFilter(normalizeValue(2), { operator: 'gt', value: '1' }), false);
  assert.equal(matchesFilter(normalizeValue(true), { operator: 'gt', value: false }), false);
});

test('matches closed between ranges only when endpoints have the cell type', () => {
  assert.equal(matchesFilter(normalizeValue(2), { operator: 'between', values: [2, 3] }), true);
  assert.equal(matchesFilter(normalizeValue(3), { operator: 'between', value: [2, 3] }), true);
  assert.equal(matchesFilter(normalizeValue(1), { operator: 'between', values: [2, 3] }), false);
  assert.equal(matchesFilter(normalizeValue(2), { operator: 'between', values: [2, '3'] }), false);
});

test('matches string operators only for strings', () => {
  assert.equal(matchesFilter(['string', 'Finance'], { operator: 'startsWith', value: 'Fin' }), true);
  assert.equal(matchesFilter(normalizeValue('spreadsheet'), { operator: 'contains', value: 'read' }), true);
  assert.equal(matchesFilter(normalizeValue('spreadsheet'), { operator: 'startsWith', value: 'spread' }), true);
  assert.equal(matchesFilter(normalizeValue('spreadsheet'), { operator: 'endsWith', value: 'sheet' }), true);
  assert.equal(matchesFilter(normalizeValue(123), { operator: 'contains', value: '2' }), false);
});

test('matches ISO date literals against date cells', () => {
  const cell = ['date', '2026-01-02T00:00:00.000Z'];
  assert.equal(
    matchesFilter(cell, { operator: 'gte', value: '2026-01-01T00:00:00.000Z' }),
    true
  );
  assert.equal(
    matchesFilter(cell, { operator: 'gte', value: '2026-01-02' }),
    true
  );
  assert.equal(
    matchesFilter(cell, { operator: 'between', values: ['2026-01-02T00:00:00.000Z', '2026-01-03T00:00:00.000Z'] }),
    true
  );
  assert.throws(() => matchesFilter(cell, { operator: 'gte', value: '2026-01-02T00:00:00' }), TypeError);
  assert.throws(() => matchesFilter(cell, { operator: 'gte', value: '2026-02-30T00:00:00Z' }), TypeError);
});

test('prints blank, scalar, formula, and object values safely', () => {
  assert.equal(printValue(['blank', null]), '');
  assert.equal(printValue(['number', 2]), '2');
  assert.equal(printValue(['formula', 'A1+1']), '"A1+1"');
  assert.equal(printValue(['formula', ['A1+1', ['number', 2]]]), '["A1+1",["number",2]]');
  assert.equal(printValue(['object', { value: 2 }]), '{"value":2}');
});
