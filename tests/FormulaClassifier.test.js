'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const { classifyFormula_, getFlattenColumnsByRow_ } = loadSources();

/**
 * @param {Record<string, {start: number, end: number}[]>} [spansBySheet]
 * @param {string[]} [affectedNamedRanges]
 * @returns {{spansBySheet: Map, affectedNamedRanges: Set}}
 */
const deletedColumnsOf = (spansBySheet = {}, affectedNamedRanges = []) => ({
  spansBySheet: new Map(Object.entries(spansBySheet)),
  affectedNamedRanges: new Set(affectedNamedRanges),
});

/**
 * Classifies a formula that lives on 'Sheet1'.
 * @param {string} formula
 * @param {{ excluded?: string[], namedRanges?: Record<string, string>, deleted?: ReturnType<typeof deletedColumnsOf> }} [options]
 * @returns {string}
 */
const classify = (formula, { excluded = [], namedRanges = {}, deleted = deletedColumnsOf() } = {}) => classifyFormula_(
  formula,
  'Sheet1',
  new Set(excluded),
  new Map(Object.entries(namedRanges)),
  deleted,
);

describe('classifyFormula_ — functions', () => {
  it('keeps an Excel-compatible formula live', () => {
    assert.equal(classify('=SUM(A1:A3)'), 'SAFE');
  });

  it('matches function names case-insensitively', () => {
    assert.equal(classify('=sum(a1:a3)'), 'SAFE');
  });

  it('flattens Sheets-only functions', () => {
    assert.equal(classify('=QUERY(A1:B2,"select *")'), 'UNSAFE');
    assert.equal(classify('=IMPORTRANGE("id","A1:B2")'), 'UNSAFE');
    assert.equal(classify('=GOOGLEFINANCE("GOOG")'), 'UNSAFE');
  });

  it('flattens spill-producing functions and INDIRECT', () => {
    assert.equal(classify('=SORT(A1:A3)'), 'UNSAFE');
    assert.equal(classify('=ARRAYFORMULA(A1:A3*2)'), 'UNSAFE');
    assert.equal(classify('=INDIRECT("A1")'), 'UNSAFE');
  });

  it('flattens a custom (unknown) function', () => {
    assert.equal(classify('=MY_CUSTOM_FN(A1)'), 'UNSAFE');
  });

  it('flattens when an unsafe function is nested inside safe ones', () => {
    assert.equal(classify('=IF(A1>0,SUM(QUERY(B1:C2,"select *")),0)'), 'UNSAFE');
  });

  it('ignores function names that only appear inside a string literal', () => {
    assert.equal(classify('=IF(A1="QUERY(x)",1,2)'), 'SAFE');
  });

  it('still sees a function that follows a string ending in a backslash', () => {
    assert.equal(classify('="C:\\"&QUERY(A1:B2,"select *")&"x"'), 'UNSAFE');
  });
});

describe('classifyFormula_ — excluded sheets', () => {
  it('flattens a reference to an excluded sheet', () => {
    assert.equal(classify('=Hidden!A1', { excluded: ['Hidden'] }), 'UNSAFE');
  });

  it('flattens a quoted reference with an escaped apostrophe', () => {
    assert.equal(classify("='Bob''s Sheet'!A1+1", { excluded: ["Bob's Sheet"] }), 'UNSAFE');
  });

  it('keeps a reference to an included sheet', () => {
    assert.equal(classify('=Data!A1', { excluded: ['Hidden'] }), 'SAFE');
  });

  it('flattens a named range that lives on an excluded sheet', () => {
    assert.equal(classify('=A1*Tax', { excluded: ['Hidden'], namedRanges: { Tax: 'Hidden' } }), 'UNSAFE');
  });

  it('keeps a named range that lives on an included sheet', () => {
    assert.equal(classify('=A1*Tax', { excluded: ['Hidden'], namedRanges: { Tax: 'Data' } }), 'SAFE');
  });
});

describe('classifyFormula_ — deleted columns', () => {
  const deleted = deletedColumnsOf({ Sheet1: [{ start: 2, end: 2 }], Other: [{ start: 0, end: 0 }] });

  it('flattens a reference to a deleted column on its own sheet', () => {
    assert.equal(classify('=C1*2', { deleted }), 'UNSAFE');
  });

  it('flattens a range that only partly overlaps a deleted column', () => {
    assert.equal(classify('=SUM(A1:D5)', { deleted }), 'UNSAFE');
    assert.equal(classify('=SUM(C:C)', { deleted }), 'UNSAFE');
    assert.equal(classify('=SUM(2:5)', { deleted }), 'UNSAFE');
  });

  it('keeps references that stay clear of the deleted column', () => {
    assert.equal(classify('=A1+B1+D1+SUM(D1:F9)', { deleted }), 'SAFE');
  });

  it('checks a qualified reference against the deleted columns of that sheet', () => {
    assert.equal(classify('=Other!A1', { deleted }), 'UNSAFE');
    assert.equal(classify('=Other!C1', { deleted }), 'SAFE');
    assert.equal(classify('=Elsewhere!C1', { deleted }), 'SAFE');
  });

  it('flattens a named range that overlaps a deleted column', () => {
    assert.equal(classify('=A1*Tax', { deleted: deletedColumnsOf({}, ['Tax']) }), 'UNSAFE');
  });
});

describe('getFlattenColumnsByRow_', () => {
  /**
   * @param {string[][]} formulas
   * @param {any[][]} values
   * @param {ReturnType<typeof deletedColumnsOf>} [deleted]
   * @returns {number[][]}
   */
  const flattenCols = (formulas, values, deleted = deletedColumnsOf()) =>
    getFlattenColumnsByRow_(formulas, values, 'Sheet1', new Set(), new Map(), deleted);

  it('selects nothing on a sheet with no unsafe formula, literals included', () => {
    assert.deepEqual(
      flattenCols([['=SUM(B1)', '', '']], [[1, 5, 'x']]),
      [[]],
    );
  });

  it('on a sheet with an unsafe formula, selects it and every non-blank formula-less cell', () => {
    assert.deepEqual(
      flattenCols(
        [['=QUERY(A1:B2,"q")', '', ''], ['', '=SUM(A1)', '']],
        [[1, 'lit', ''], ['a', 2, null]],
      ),
      [[0, 1], [0]],
    );
  });

  it('leaves a safe live formula alone even on a sheet with unsafe ones', () => {
    assert.deepEqual(
      flattenCols([['=QUERY(A1:B2,"q")', '=SUM(A1)']], [[1, 2]]),
      [[0]],
    );
  });

  it('treats a literal 0 or false as non-blank', () => {
    assert.deepEqual(
      flattenCols([['=MY_FN(1)', '', '']], [[1, 0, false]]),
      [[0, 1, 2]],
    );
  });

  it('never selects a cell in a deleted column', () => {
    const deleted = deletedColumnsOf({ Sheet1: [{ start: 1, end: 1 }] });
    assert.deepEqual(
      flattenCols([['=MY_FN(1)', '=MY_FN(2)', '']], [[1, 2, 'y']], deleted),
      [[0, 2]],
    );
  });

  it('counts a formula in a deleted column as unsafe, so the rest of the sheet is rewritten', () => {
    const deleted = deletedColumnsOf({ Sheet1: [{ start: 1, end: 1 }] });
    assert.deepEqual(
      flattenCols([['', '=A2', '']], [['x', 5, 'y']], deleted),
      [[0, 2]],
    );
  });

  it('does not rewrite literals when the deleted column holds only a literal', () => {
    const deleted = deletedColumnsOf({ Sheet1: [{ start: 1, end: 1 }] });
    assert.deepEqual(
      flattenCols([['', '']], [['x', 'y']], deleted),
      [[]],
    );
  });

  it('flattens a formula that reads a deleted column', () => {
    const deleted = deletedColumnsOf({ Sheet1: [{ start: 2, end: 2 }] });
    assert.deepEqual(
      flattenCols([['=SUM(C1:C3)', '', '']], [[6, 'lit', 1]], deleted),
      [[0, 1]],
    );
  });
});
