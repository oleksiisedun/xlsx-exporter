'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const { classifyFormula_ } = loadSources();

/**
 * @param {string} formula
 * @returns {string}
 */
const classify = (formula) => classifyFormula_(
  formula,
  'Sheet1',
  new Set(),
  new Map(),
  { spansBySheet: new Map(), affectedNamedRanges: new Set() },
);

describe('classifyFormula_', () => {
  it('keeps an Excel-compatible formula live', () => {
    assert.equal(classify('=SUM(A1:A3)'), 'SAFE');
  });

  it('flattens a Sheets-only function', () => {
    assert.equal(classify('=QUERY(A1:B2,"select *")'), 'UNSAFE');
  });

  it('ignores function names that only appear inside a string literal', () => {
    assert.equal(classify('=IF(A1="QUERY(x)",1,2)'), 'SAFE');
  });

  it('still sees a function that follows a string ending in a backslash', () => {
    assert.equal(classify('="C:\\"&QUERY(A1:B2,"select *")&"x"'), 'UNSAFE');
  });
});
