'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const { columnLettersToIndex_, spansOverlap_ } = loadSources();

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

describe('spansOverlap_', () => {
  const span = (start, end) => ({ start, end });

  it('is true when a range overlaps a span', () => {
    assert.equal(spansOverlap_([span(2, 5)], 4, 8), true);
  });

  it('is true when a range is fully contained within a span', () => {
    assert.equal(spansOverlap_([span(0, 10)], 4, 5), true);
  });

  it('is false when a range falls entirely outside every span', () => {
    assert.equal(spansOverlap_([span(2, 5), span(10, 12)], 6, 9), false);
  });

  it('is false for an empty span list', () => {
    assert.equal(spansOverlap_([], 0, 100), false);
  });

  it('is true at the exact boundary of a span', () => {
    assert.equal(spansOverlap_([span(2, 5)], 5, 7), true);
    assert.equal(spansOverlap_([span(2, 5)], 0, 2), true);
  });
});
