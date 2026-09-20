'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const { resolveIncludedSheetNames_ } = loadSources();

const ALL = ['Summary', 'Data', 'Notes'];

describe('resolveIncludedSheetNames_', () => {
  it('includes every sheet when no option is given', () => {
    assert.deepEqual(resolveIncludedSheetNames_(ALL, undefined, undefined), ALL);
  });

  it('treats empty arrays as "not given"', () => {
    assert.deepEqual(resolveIncludedSheetNames_(ALL, [], []), ALL);
    assert.deepEqual(resolveIncludedSheetNames_(ALL, [], ['Notes']), ['Summary', 'Data']);
  });

  it("keeps the source's tab order regardless of the order in includeSheets", () => {
    assert.deepEqual(resolveIncludedSheetNames_(ALL, ['Notes', 'Summary'], undefined), ['Summary', 'Notes']);
  });

  it('removes excluded sheets', () => {
    assert.deepEqual(resolveIncludedSheetNames_(ALL, undefined, ['Data']), ['Summary', 'Notes']);
  });

  it('rejects includeSheets and excludeSheets together', () => {
    assert.throws(
      () => resolveIncludedSheetNames_(ALL, ['Data'], ['Notes']),
      /either includeSheets or excludeSheets, not both/,
    );
  });

  it('rejects unknown sheet names, listing them', () => {
    assert.throws(
      () => resolveIncludedSheetNames_(ALL, ['Data', 'Nope', 'Missing'], undefined),
      /includeSheets contains unknown sheet name\(s\): Nope, Missing/,
    );
    assert.throws(
      () => resolveIncludedSheetNames_(ALL, undefined, ['Nope']),
      /excludeSheets contains unknown sheet name\(s\): Nope/,
    );
  });

  it('rejects an export with no sheets left', () => {
    assert.throws(
      () => resolveIncludedSheetNames_(ALL, undefined, ALL),
      /included-sheet set is empty/,
    );
  });
});
