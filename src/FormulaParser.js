/**
 * Removes double-quoted string literals from a formula so their contents
 * aren't mistaken for function names or sheet references. Sheets escapes a
 * quote inside a string by doubling it (`""`); a backslash is an ordinary
 * character, so `"C:\"` is a complete string.
 * @param {string} formula
 * @returns {string}
 */
function stripStringLiterals_(formula) {
  return formula.replace(/"(?:[^"]|"")*"/g, '""');
}

/**
 * Extracts the set of function names called directly in a formula.
 * @param {string} formulaWithoutStrings
 * @returns {string[]} Upper-cased, de-duplicated function names.
 */
function extractFunctionNames_(formulaWithoutStrings) {
  const names = new Set();
  const re = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;
  let m;
  while ((m = re.exec(formulaWithoutStrings))) names.add(m[1].toUpperCase());
  return [...names];
}

/** Matches SheetName! or 'Sheet Name'! (with '' as the escaped-quote sequence). */
const SHEET_REF_RE = /'((?:[^']|'')*)'!|([A-Za-z_][A-Za-z0-9_.]*)!/g;

/**
 * Unescapes a quoted sheet name's doubled single quotes ('' -> ').
 * @param {string} quotedName
 * @returns {string}
 */
function unescapeSheetName_(quotedName) {
  return quotedName.replace(/''/g, "'");
}

/**
 * Extracts sheet names referenced via SheetName!A1 / 'Sheet Name'!A1:B2 syntax.
 * @param {string} formulaWithoutStrings
 * @returns {string[]} Unescaped sheet names (de-duplicated).
 */
function extractSheetQualifiedReferences_(formulaWithoutStrings) {
  const names = new Set();
  let m;
  SHEET_REF_RE.lastIndex = 0;
  while ((m = SHEET_REF_RE.exec(formulaWithoutStrings))) {
    const raw = m[1] !== undefined ? unescapeSheetName_(m[1]) : m[2];
    names.add(raw);
  }
  return [...names];
}

/**
 * Extracts bare identifier tokens that could be named-range references
 * (i.e. not function calls, not part of a sheet-qualified reference, not a
 * plain cell address like A1, not TRUE/FALSE).
 *
 * Known limitations: a sheet name embedded inside a text/QUERY-criteria
 * string won't be caught by this regex-based approach, but QUERY is always
 * unsafe via the blocklist so this is moot in practice. INDIRECT's string
 * argument is never parsed, even if it's a literal — INDIRECT is always
 * flattened, so it's safe by construction. Only spreadsheet-level named
 * ranges are considered (Range.getFormulas() always returns A1-style text,
 * so R1C1 notation isn't a concern).
 * @param {string} formulaWithoutStrings
 * @returns {string[]}
 */
function extractBareIdentifiers_(formulaWithoutStrings) {
  const withoutSheetRefs = formulaWithoutStrings.replace(SHEET_REF_RE, '');
  const found = new Set();
  const re = /\b([A-Za-z_][A-Za-z0-9_.]*)\b(?!\s*\()/g;
  let m;
  while ((m = re.exec(withoutSheetRefs))) {
    const token = m[1];
    if (/^[A-Za-z]{1,3}\d+$/.test(token)) continue; // plain cell reference, e.g. A1
    const upper = token.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') continue;
    found.add(token);
  }
  return [...found];
}

/**
 * Matches a cell/range reference with an optional sheet qualifier, in four
 * shapes: `A1` / `A1:B2` / `A1:B` (groups `cellCol`, `rangeEndCol`), a
 * whole-column range `A:C` (`colRangeStart`, `colRangeEnd`), and a whole-row
 * range `2:5` (`rowRange`). `$` anchors are accepted. The lookbehind/lookahead
 * keep it from matching inside a longer identifier or a function call such as
 * `LOG10(`; a 1-3 letter token with no digits is only a reference in the
 * `A:C` shape, so bare named ranges like `Tax` don't match.
 */
const CELL_REFERENCE_RE = new RegExp(
  "(?<![A-Za-z0-9_.$!'])" +
  "(?:(?:'(?<quotedSheet>(?:[^']|'')*)'|(?<plainSheet>[A-Za-z_][A-Za-z0-9_.]*))!)?" +
  '(?:' +
    '\\$?(?<cellCol>[A-Za-z]{1,3})\\$?\\d+(?::\\$?(?<rangeEndCol>[A-Za-z]{1,3})\\$?\\d*)?' +
    '|\\$?(?<colRangeStart>[A-Za-z]{1,3}):\\$?(?<colRangeEnd>[A-Za-z]{1,3})' +
    '|(?<rowRange>\\$?\\d+:\\$?\\d+)' +
  ')' +
  '(?![A-Za-z0-9_(.])',
  'g'
);

/**
 * @typedef {Object} ColumnReference
 * @property {string|null} sheetName - Sheet the reference is qualified with, or null if unqualified (i.e. the formula's own sheet).
 * @property {number} startCol - 0-based first column covered.
 * @property {number} endCol - 0-based last column covered; Infinity for a whole-row reference (`2:5`), which spans every column.
 */

/**
 * Extracts the columns touched by every cell/range reference in a formula.
 * Used to detect formulas that read a column which will be deleted from the
 * export. Errs toward over-matching (a false positive only flattens an
 * otherwise-safe formula to its value).
 * @param {string} formulaWithoutStrings
 * @returns {ColumnReference[]}
 */
function extractColumnReferences_(formulaWithoutStrings) {
  /** @type {ColumnReference[]} */
  const references = [];
  for (const m of formulaWithoutStrings.matchAll(CELL_REFERENCE_RE)) {
    const g = m.groups ?? {};
    const sheetName = g.quotedSheet !== undefined ? unescapeSheetName_(g.quotedSheet) : g.plainSheet ?? null;
    if (g.rowRange) {
      references.push({ sheetName, startCol: 0, endCol: Infinity });
      continue;
    }
    const firstCol = columnLettersToIndex_(g.cellCol ?? g.colRangeStart);
    const lastCol = columnLettersToIndex_(g.rangeEndCol ?? g.colRangeEnd ?? g.cellCol);
    references.push({ sheetName, startCol: Math.min(firstCol, lastCol), endCol: Math.max(firstCol, lastCol) });
  }
  return references;
}
