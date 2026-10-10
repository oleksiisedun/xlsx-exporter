'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const { watchedValuesEqual_ } = loadSources();

describe('watchedValuesEqual_', () => {
  it('is true for identical snapshots, including empty ones', () => {
    assert.equal(watchedValuesEqual_(new Map(), new Map()), true);
    assert.equal(watchedValuesEqual_(new Map([['A', []]]), new Map([['A', []]])), true);
    assert.equal(watchedValuesEqual_(new Map([['A', [1, 'x', null]]]), new Map([['A', [1, 'x', null]]])), true);
  });

  it('is false when any single value differs', () => {
    assert.equal(watchedValuesEqual_(new Map([['A', [1, 2]]]), new Map([['A', [1, 3]]])), false);
  });

  it('does not conflate 0 with an empty string or a different type', () => {
    assert.equal(watchedValuesEqual_(new Map([['A', [0]]]), new Map([['A', ['']]])), false);
    assert.equal(watchedValuesEqual_(new Map([['A', [1]]]), new Map([['A', ['1']]])), false);
  });

  it('is false when the number of watched cells differs', () => {
    assert.equal(watchedValuesEqual_(new Map([['A', [1]]]), new Map([['A', [1, 2]]])), false);
  });

  it('is false when a sheet is missing from the second snapshot', () => {
    assert.equal(watchedValuesEqual_(new Map([['A', [1]]]), new Map([['B', [1]]])), false);
  });

  it('compares dates by timestamp, since getValues() returns a new Date each call', () => {
    const t = Date.UTC(2026, 0, 1);
    assert.equal(watchedValuesEqual_(new Map([['A', [new Date(t)]]]), new Map([['A', [new Date(t)]]])), true);
    assert.equal(watchedValuesEqual_(new Map([['A', [new Date(t)]]]), new Map([['A', [new Date(t + 1)]]])), false);
  });

  it('is false when only one side is a date', () => {
    const t = Date.UTC(2026, 0, 1);
    assert.equal(watchedValuesEqual_(new Map([['A', [new Date(t)]]]), new Map([['A', [t]]])), false);
  });
});

describe('readFlattenCellsSnapshot_', () => {
  const { readFlattenCellsSnapshot_ } = loadSources();
  const noDeleted = { spansBySheet: new Map(), affectedNamedRanges: new Set() };

  /** Stub exposing only what the function calls; the grid is swapped between polls. */
  const stubSpreadsheet = (/** @type {{formulas: string[][], values: any[][]}} */ grid) => ({
    getSheetByName: () => ({
      getDataRange: () => ({
        getNumRows: () => grid.values.length,
        getNumColumns: () => grid.values[0].length,
        getFormulas: () => grid.formulas,
        getValues: () => grid.values,
      }),
      getRange: (/** @type {number} */ r, /** @type {number} */ c) => ({ getA1Notation: () => `R${r}C${c}` }),
    }),
  });
  const read = (/** @type {any} */ grid) => readFlattenCellsSnapshot_(stubSpreadsheet(grid), ['S'], new Set(), new Map(), noDeleted);
  const importFormula = '=IMPORTRANGE("id","A1:B2")';

  it('picks up spill cells that appear after the anchor stops loading', () => {
    const loading = read({ formulas: [[importFormula, ''], ['', '']], values: [['Loading...', ''], ['', '']] });
    assert.deepEqual(loading.positions.get('S'), ['0:0']);
    assert.deepEqual(loading.loadingCell, { sheetName: 'S', a1Notation: 'R1C1' });

    const settled = read({ formulas: [[importFormula, ''], ['', '']], values: [[0, 2], [0, 6]] });
    assert.deepEqual(settled.positions.get('S'), ['0:0', '0:1', '1:0', '1:1']);
    assert.equal(settled.loadingCell, null);
  });

  it('treats a spill that grows as a change even when the anchor value is unchanged', () => {
    const stale = read({ formulas: [[importFormula, ''], ['', '']], values: [[0, ''], ['', '']] });
    const real = read({ formulas: [[importFormula, ''], ['', '']], values: [[0, 2], ['', '']] });
    assert.equal(watchedValuesEqual_(stale.positions, real.positions), false);
  });
});
