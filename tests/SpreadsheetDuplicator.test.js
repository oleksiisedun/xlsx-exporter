'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const { mergeColumnsIntoRectangles_ } = loadSources();

const rect = (row, col, numRows, numCols) => ({ row, col, numRows, numCols });

describe('mergeColumnsIntoRectangles_', () => {
  it('returns nothing when no cells are selected', () => {
    assert.deepEqual(mergeColumnsIntoRectangles_([]), []);
    assert.deepEqual(mergeColumnsIntoRectangles_([[], []]), []);
  });

  it('turns a contiguous run into one rectangle', () => {
    assert.deepEqual(mergeColumnsIntoRectangles_([[1, 2, 3]]), [rect(0, 1, 1, 3)]);
  });

  it('splits non-contiguous columns into separate rectangles', () => {
    assert.deepEqual(mergeColumnsIntoRectangles_([[0, 2]]), [rect(0, 0, 1, 1), rect(0, 2, 1, 1)]);
  });

  it('stacks identical runs on consecutive rows', () => {
    assert.deepEqual(mergeColumnsIntoRectangles_([[0, 1], [0, 1], [0, 1]]), [rect(0, 0, 3, 2)]);
  });

  it('stacks each run independently when a row has several', () => {
    assert.deepEqual(
      mergeColumnsIntoRectangles_([[0, 2], [0, 2]]),
      [rect(0, 0, 2, 1), rect(0, 2, 2, 1)],
    );
  });

  it('starts a new rectangle after an empty row', () => {
    assert.deepEqual(
      mergeColumnsIntoRectangles_([[0, 1], [], [0, 1]]),
      [rect(0, 0, 1, 2), rect(2, 0, 1, 2)],
    );
  });

  it('does not stack runs whose column extent differs', () => {
    assert.deepEqual(mergeColumnsIntoRectangles_([[0, 1], [0]]), [rect(0, 0, 1, 2), rect(1, 0, 1, 1)]);
    assert.deepEqual(mergeColumnsIntoRectangles_([[0], [1]]), [rect(0, 0, 1, 1), rect(1, 1, 1, 1)]);
  });

  it('covers exactly the selected cells', () => {
    const selection = [[0, 1], [0, 1], [], [0, 1], [3], [3, 4, 5], [0, 2, 3]];
    const covered = mergeColumnsIntoRectangles_(selection).flatMap(({ row, col, numRows, numCols }) =>
      Array.from({ length: numRows * numCols }, (_, i) => `${row + Math.floor(i / numCols)},${col + (i % numCols)}`));
    const expected = selection.flatMap((cols, r) => cols.map((c) => `${r},${c}`));
    assert.deepEqual(covered.sort(), expected.sort());
  });
});
