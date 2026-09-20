'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const { columnLettersToIndex_ } = loadSources();

describe('columnLettersToIndex_', () => {
  it('converts single letters', () => {
    assert.equal(columnLettersToIndex_('A'), 0);
    assert.equal(columnLettersToIndex_('Z'), 25);
  });

  it('carries into two and three letters', () => {
    assert.equal(columnLettersToIndex_('AA'), 26);
    assert.equal(columnLettersToIndex_('AZ'), 51);
    assert.equal(columnLettersToIndex_('BA'), 52);
    assert.equal(columnLettersToIndex_('XFD'), 16383);
  });

  it('is case-insensitive', () => {
    assert.equal(columnLettersToIndex_('ab'), columnLettersToIndex_('AB'));
  });
});
