'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const {
  stripStringLiterals_,
  extractFunctionNames_,
  extractSheetQualifiedReferences_,
  extractBareIdentifiers_,
  extractColumnReferences_,
} = loadSources();

describe('stripStringLiterals_', () => {
  it('empties a plain string literal', () => {
    assert.equal(stripStringLiterals_('=IF(A1="QUERY(x)",1,2)'), '=IF(A1="",1,2)');
  });

  it('treats a doubled quote as an escaped quote inside the string', () => {
    assert.equal(stripStringLiterals_('="say ""QUERY("" now"&A1'), '=""&A1');
  });

  it('does not treat a backslash as an escape (Sheets has no backslash escapes)', () => {
    // The string is just `C:\`; QUERY( sits outside it and must stay visible.
    assert.equal(stripStringLiterals_('="C:\\"&QUERY(A1:B2,"q")'), '=""&QUERY(A1:B2,"")');
  });
});

describe('extractFunctionNames_', () => {
  it('finds upper-cased, de-duplicated names, including dotted and spaced calls', () => {
    assert.deepEqual(
      extractFunctionNames_('=SUM(A1)+ceiling.math(2)+my_fn (3)+SUM(B1)'),
      ['SUM', 'CEILING.MATH', 'MY_FN'],
    );
  });

  it('finds nested calls', () => {
    assert.deepEqual(extractFunctionNames_('=IF(ISNUMBER(A1),ROUND(A1,2),0)'), ['IF', 'ISNUMBER', 'ROUND']);
  });

  it('returns nothing for a formula with no calls', () => {
    assert.deepEqual(extractFunctionNames_('=A1+B1*2'), []);
  });
});

describe('extractSheetQualifiedReferences_', () => {
  it('finds plain and quoted sheet names', () => {
    assert.deepEqual(extractSheetQualifiedReferences_("=Data!B2+'My Sheet'!A1:B2"), ['Data', 'My Sheet']);
  });

  it("unescapes doubled apostrophes in quoted names", () => {
    assert.deepEqual(extractSheetQualifiedReferences_("='Bob''s'!A1"), ["Bob's"]);
  });

  it('de-duplicates repeated references', () => {
    assert.deepEqual(extractSheetQualifiedReferences_('=Data!A1+Data!B1'), ['Data']);
  });

  it('returns nothing for unqualified references', () => {
    assert.deepEqual(extractSheetQualifiedReferences_('=A1+B2'), []);
  });

  it('gives the same answer when called repeatedly (shared global regex state)', () => {
    extractSheetQualifiedReferences_('=Data!A1');
    assert.deepEqual(extractSheetQualifiedReferences_('=Other!A1'), ['Other']);
  });
});

describe('extractBareIdentifiers_', () => {
  it('finds a named range', () => {
    assert.deepEqual(extractBareIdentifiers_('=A1*Tax'), ['Tax']);
  });

  it('skips function names, cell addresses and booleans', () => {
    assert.deepEqual(extractBareIdentifiers_('=IF(TRUE,SUM(A1:B2),false)'), []);
  });

  it('skips sheet-qualified references', () => {
    assert.deepEqual(extractBareIdentifiers_("=Data!A1+'My Sheet'!B2"), []);
  });

  it('keeps a named range used as a function argument', () => {
    assert.deepEqual(extractBareIdentifiers_('=SUM(Tax,Fee)'), ['Tax', 'Fee']);
  });

  it('reports the letters of $-anchored references as identifiers (harmless unless a named range is called "A")', () => {
    assert.deepEqual(extractBareIdentifiers_('=SUM($A$1:$B$2)'), ['A', 'B']);
  });
});

describe('extractColumnReferences_', () => {
  const ref = (sheetName, startCol, endCol) => ({ sheetName, startCol, endCol });

  it('reads a single cell', () => {
    assert.deepEqual(extractColumnReferences_('=A1'), [ref(null, 0, 0)]);
  });

  it('reads a range and ignores $ anchors', () => {
    assert.deepEqual(extractColumnReferences_('=SUM($A$1:$C2)'), [ref(null, 0, 2)]);
  });

  it('reads an open-ended range like A1:B', () => {
    assert.deepEqual(extractColumnReferences_('=SUM(A1:B)'), [ref(null, 0, 1)]);
  });

  it('normalizes a reversed range', () => {
    assert.deepEqual(extractColumnReferences_('=SUM(C1:A1)'), [ref(null, 0, 2)]);
  });

  it('reads a whole-column range', () => {
    assert.deepEqual(extractColumnReferences_('=SUM(B:D)'), [ref(null, 1, 3)]);
  });

  it('treats a whole-row range as spanning every column', () => {
    assert.deepEqual(extractColumnReferences_('=SUM(2:5)'), [ref(null, 0, Infinity)]);
  });

  it('captures plain and quoted sheet qualifiers', () => {
    assert.deepEqual(
      extractColumnReferences_("=Data!C3+'Bob''s Sheet'!A:B"),
      [ref('Data', 2, 2), ref("Bob's Sheet", 0, 1)],
    );
  });

  it('reads multi-letter columns', () => {
    assert.deepEqual(extractColumnReferences_('=AA1+XFD9'), [ref(null, 26, 26), ref(null, 16383, 16383)]);
  });

  it('does not mistake function calls or named ranges for references', () => {
    assert.deepEqual(extractColumnReferences_('=LOG10(3)+Tax+Rate_2023'), []);
  });

  it('does treat a named-range-looking token that is a valid cell address as a reference', () => {
    assert.deepEqual(extractColumnReferences_('=Fee2023'), [ref(null, 4190, 4190)]);
  });

  it('finds every reference in a formula', () => {
    assert.deepEqual(extractColumnReferences_('=A1+SUM(C1:D9)'), [ref(null, 0, 0), ref(null, 2, 3)]);
  });
});
