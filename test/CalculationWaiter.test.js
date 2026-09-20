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
