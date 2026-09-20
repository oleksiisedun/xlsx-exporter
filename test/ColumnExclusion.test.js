'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const {
  resolveExcludeColumnsMode_,
  parseColumnSpec_,
  mergeColumnSpans_,
  resolveExcludedColumnSpans_,
  buildDeletedColumns_,
} = loadSources();

const span = (start, end) => ({ start, end });

/**
 * Minimal stand-in for a Spreadsheet: only what the resolver reads.
 * @param {Record<string, number>} maxColumnsBySheet
 * @returns {object}
 */
const spreadsheetWithSheets = (maxColumnsBySheet) => ({
  getName: () => 'Test',
  getSheetByName: (name) => (name in maxColumnsBySheet ? { getMaxColumns: () => maxColumnsBySheet[name] } : null),
});

describe('resolveExcludeColumnsMode_', () => {
  it("defaults to 'delete'", () => {
    assert.equal(resolveExcludeColumnsMode_(undefined), 'delete');
  });

  it("accepts 'delete' and 'hide'", () => {
    assert.equal(resolveExcludeColumnsMode_('delete'), 'delete');
    assert.equal(resolveExcludeColumnsMode_('hide'), 'hide');
  });

  it('rejects anything else, naming the bad value', () => {
    assert.throws(() => resolveExcludeColumnsMode_('Hide'), /excludeColumnsMode must be 'delete' or 'hide', got "Hide"/);
    assert.throws(() => resolveExcludeColumnsMode_(''), /excludeColumnsMode/);
  });
});

describe('parseColumnSpec_', () => {
  it('parses a single column, case-insensitively and with a $ anchor', () => {
    assert.deepEqual(parseColumnSpec_('C', 'Data'), span(2, 2));
    assert.deepEqual(parseColumnSpec_('c', 'Data'), span(2, 2));
    assert.deepEqual(parseColumnSpec_('$C', 'Data'), span(2, 2));
  });

  it('parses a range and normalizes a reversed one', () => {
    assert.deepEqual(parseColumnSpec_('F:H', 'Data'), span(5, 7));
    assert.deepEqual(parseColumnSpec_('H:F', 'Data'), span(5, 7));
    assert.deepEqual(parseColumnSpec_('$F:$H', 'Data'), span(5, 7));
  });

  it('parses multi-letter columns and tolerates surrounding whitespace', () => {
    assert.deepEqual(parseColumnSpec_(' AA ', 'Data'), span(26, 26));
  });

  it('rejects malformed specs, naming the sheet and the spec', () => {
    for (const bad of ['', '1', 'C1', 'A:B:C', 'ABCD', 'A-B', 42, null]) {
      assert.throws(() => parseColumnSpec_(bad, 'Data'), /excludeColumns\["Data"\] contains invalid column/, String(bad));
    }
  });
});

describe('mergeColumnSpans_', () => {
  it('merges overlapping spans', () => {
    assert.deepEqual(mergeColumnSpans_([span(0, 2), span(1, 4)]), [span(0, 4)]);
  });

  it('merges spans that merely touch', () => {
    assert.deepEqual(mergeColumnSpans_([span(0, 1), span(2, 3)]), [span(0, 3)]);
  });

  it('keeps spans with a gap apart', () => {
    assert.deepEqual(mergeColumnSpans_([span(0, 1), span(3, 4)]), [span(0, 1), span(3, 4)]);
  });

  it('sorts unsorted input and absorbs a contained span', () => {
    assert.deepEqual(mergeColumnSpans_([span(7, 7), span(1, 5), span(2, 3)]), [span(1, 5), span(7, 7)]);
  });

  it('handles an empty list', () => {
    assert.deepEqual(mergeColumnSpans_([]), []);
  });

  it('does not mutate its input', () => {
    const input = [span(0, 1), span(1, 4)];
    mergeColumnSpans_(input);
    assert.deepEqual(input, [span(0, 1), span(1, 4)]);
  });
});

describe('resolveExcludedColumnSpans_', () => {
  const sheets = spreadsheetWithSheets({ Data: 10, Other: 5 });

  it('returns an empty map when nothing is excluded', () => {
    assert.equal(resolveExcludedColumnSpans_(sheets, ['Data'], undefined, 'delete').size, 0);
    assert.equal(resolveExcludedColumnSpans_(sheets, ['Data'], {}, 'delete').size, 0);
  });

  it('resolves and merges specs per sheet', () => {
    const result = resolveExcludedColumnSpans_(sheets, ['Data', 'Other'], { Data: ['C', 'B:D', 'H'], Other: ['a'] }, 'delete');
    assert.deepEqual(result, new Map([['Data', [span(1, 3), span(7, 7)]], ['Other', [span(0, 0)]]]));
  });

  it("clips a span to the sheet's grid", () => {
    const result = resolveExcludedColumnSpans_(sheets, ['Data'], { Data: ['H:Z'] }, 'delete');
    assert.deepEqual(result.get('Data'), [span(7, 9)]);
  });

  it('drops a span entirely beyond the grid, and the sheet with it', () => {
    const result = resolveExcludedColumnSpans_(sheets, ['Data'], { Data: ['Z'] }, 'delete');
    assert.equal(result.size, 0);
  });

  it('rejects a sheet that is not part of the export', () => {
    assert.throws(
      () => resolveExcludedColumnSpans_(sheets, ['Data'], { Other: ['A'] }, 'delete'),
      /names sheet "Other", which is not part of the export/,
    );
  });

  it('rejects specs that are not an array', () => {
    assert.throws(
      () => resolveExcludedColumnSpans_(sheets, ['Data'], { Data: 'C' }, 'delete'),
      /excludeColumns\["Data"\] must be an array/,
    );
  });

  it('rejects a malformed spec before doing anything else', () => {
    assert.throws(
      () => resolveExcludedColumnSpans_(sheets, ['Data'], { Data: ['C', '3'] }, 'delete'),
      /invalid column "3"/,
    );
  });

  it('refuses to delete every column of a sheet, even when built from touching specs', () => {
    assert.throws(
      () => resolveExcludedColumnSpans_(sheets, ['Data'], { Data: ['A:E', 'F:J'] }, 'delete'),
      /would delete every column of "Data"/,
    );
    assert.throws(
      () => resolveExcludedColumnSpans_(sheets, ['Data'], { Data: ['A:Z'] }, 'delete'),
      /would delete every column of "Data"/,
    );
  });

  it("allows covering every column in 'hide' mode", () => {
    const result = resolveExcludedColumnSpans_(sheets, ['Data'], { Data: ['A:Z'] }, 'hide');
    assert.deepEqual(result.get('Data'), [span(0, 9)]);
  });

  it('allows deleting all but one column', () => {
    const result = resolveExcludedColumnSpans_(sheets, ['Data'], { Data: ['B:J'] }, 'delete');
    assert.deepEqual(result.get('Data'), [span(1, 9)]);
  });
});

describe('buildDeletedColumns_', () => {
  /**
   * @param {string} name
   * @param {string} sheetName
   * @param {number} firstColumn - 1-based, as Range.getColumn() returns.
   * @param {number} numColumns
   * @returns {object}
   */
  const namedRange = (name, sheetName, firstColumn, numColumns) => ({
    getName: () => name,
    getRange: () => ({
      getSheet: () => ({ getName: () => sheetName }),
      getColumn: () => firstColumn,
      getNumColumns: () => numColumns,
    }),
  });
  const spreadsheetWithNamedRanges = (...ranges) => ({ getNamedRanges: () => ranges });
  const spans = new Map([['Data', [span(2, 3)]]]);

  it("is empty in 'hide' mode, since hidden columns keep their data", () => {
    const result = buildDeletedColumns_(spreadsheetWithNamedRanges(namedRange('Tax', 'Data', 3, 1)), spans, 'hide');
    assert.equal(result.spansBySheet.size, 0);
    assert.equal(result.affectedNamedRanges.size, 0);
  });

  it('is empty when no columns are excluded', () => {
    const result = buildDeletedColumns_(spreadsheetWithNamedRanges(), new Map(), 'delete');
    assert.equal(result.spansBySheet.size, 0);
  });

  it('exposes the spans and flags only named ranges that overlap them', () => {
    const result = buildDeletedColumns_(
      spreadsheetWithNamedRanges(
        namedRange('Inside', 'Data', 3, 1), // column C
        namedRange('Straddles', 'Data', 2, 2), // B:C
        namedRange('Before', 'Data', 1, 2), // A:B
        namedRange('After', 'Data', 5, 2), // E:F
        namedRange('OtherSheet', 'Other', 3, 1), // C, but on a sheet with no deletions
      ),
      spans,
      'delete',
    );
    assert.equal(result.spansBySheet, spans);
    assert.deepEqual([...result.affectedNamedRanges].sort(), ['Inside', 'Straddles']);
  });
});
