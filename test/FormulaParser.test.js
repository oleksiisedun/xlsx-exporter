'use strict';

const { describe, it } = require('node:test');
const assert = require('node:assert/strict');
const { loadSources } = require('./loadSources');

const { stripStringLiterals_ } = loadSources();

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
