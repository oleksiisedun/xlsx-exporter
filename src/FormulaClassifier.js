/** Google-Sheets-only functions with no Excel equivalent. */
const UNSAFE_SHEETS_ONLY_FUNCTIONS = new Set([
  'IMPORTRANGE', 'GOOGLEFINANCE', 'GOOGLETRANSLATE', 'DETECTLANGUAGE',
  'IMPORTHTML', 'IMPORTXML', 'IMPORTDATA', 'IMPORTFEED', 'QUERY', 'SPARKLINE',
]);

/**
 * Functions that DO have an Excel equivalent but are force-flattened anyway:
 * - INDIRECT resolves a reference from a string at runtime; its target sheet
 *   can never be statically verified, so it's always flattened.
 * - SORT/UNIQUE/FILTER/SEQUENCE/RANDARRAY/ARRAYFORMULA are spill-producing;
 *   flattening avoids any spill/#REF collision when the duplicate recalculates.
 */
const ALWAYS_FLATTEN_FUNCTIONS = new Set([
  'INDIRECT', 'SORT', 'UNIQUE', 'FILTER', 'SEQUENCE', 'RANDARRAY', 'ARRAYFORMULA',
]);

/**
 * Allowlist gate: standard functions shared identically (in name + core
 * semantics) between Google Sheets and Excel. Anything called in a formula
 * that is NOT in this set is treated as unsafe — this is what catches custom
 * Apps Script functions and any Sheets-only function not explicitly
 * blocklisted above. Extend this list as false positives are found during
 * testing; keep it as the single source of truth (do not duplicate elsewhere).
 */
const KNOWN_EXCEL_COMPATIBLE_FUNCTIONS = new Set([
  // Math
  'SUM', 'SUMIF', 'SUMIFS', 'SUMPRODUCT', 'SUMSQ', 'PRODUCT', 'ABS', 'ROUND', 'ROUNDUP',
  'ROUNDDOWN', 'MROUND', 'CEILING', 'FLOOR', 'INT', 'TRUNC', 'MOD', 'POWER', 'SQRT', 'EXP',
  'LN', 'LOG', 'LOG10', 'PI', 'RAND', 'RANDBETWEEN', 'SIGN', 'GCD', 'LCM', 'FACT', 'COMBIN',
  'PERMUT', 'DEGREES', 'RADIANS', 'SIN', 'COS', 'TAN', 'ASIN', 'ACOS', 'ATAN', 'ATAN2',
  'SINH', 'COSH', 'TANH', 'SUBTOTAL', 'AGGREGATE', 'CEILING.MATH', 'FLOOR.MATH',
  // Statistical
  'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTBLANK',
  'COUNTIF', 'COUNTIFS', 'MAX', 'MAXA', 'MAXIFS', 'MIN', 'MINA', 'MINIFS', 'MEDIAN',
  'MODE', 'STDEV', 'STDEVA', 'STDEVP', 'STDEVPA', 'VAR', 'VARA', 'VARP', 'VARPA', 'LARGE',
  'SMALL', 'RANK', 'PERCENTILE', 'QUARTILE', 'TRIMMEAN', 'CORREL', 'COVAR', 'FORECAST',
  'TREND', 'GROWTH', 'SLOPE', 'INTERCEPT', 'RSQ', 'NORMDIST', 'NORMSDIST', 'NORMINV',
  'NORMSINV', 'STDEV.S', 'STDEV.P', 'VAR.S', 'VAR.P', 'MODE.SNGL', 'PERCENTILE.INC',
  'PERCENTILE.EXC', 'QUARTILE.INC', 'QUARTILE.EXC', 'RANK.EQ', 'RANK.AVG', 'NORM.DIST',
  'NORM.S.DIST', 'NORM.INV', 'NORM.S.INV', 'COVARIANCE.P', 'COVARIANCE.S',
  // Logical
  'IF', 'IFS', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT', 'XOR', 'TRUE', 'FALSE', 'SWITCH',
  // Text
  'CONCATENATE', 'CONCAT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'LOWER', 'UPPER',
  'PROPER', 'TRIM', 'SUBSTITUTE', 'REPLACE', 'FIND', 'SEARCH', 'TEXT', 'VALUE', 'REPT',
  'CHAR', 'CODE', 'EXACT', 'CLEAN', 'DOLLAR', 'FIXED', 'UNICHAR', 'UNICODE',
  // Date/Time
  'DATE', 'DATEVALUE', 'TIME', 'TIMEVALUE', 'NOW', 'TODAY', 'YEAR', 'MONTH', 'DAY',
  'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'WEEKNUM', 'DAYS', 'DAYS360', 'NETWORKDAYS',
  'WORKDAY', 'EDATE', 'EOMONTH', 'DATEDIF', 'ISOWEEKNUM', 'NETWORKDAYS.INTL',
  'WORKDAY.INTL',
  // Lookup/Reference
  'VLOOKUP', 'HLOOKUP', 'LOOKUP', 'INDEX', 'MATCH', 'XLOOKUP', 'XMATCH', 'CHOOSE', 'ROW',
  'ROWS', 'COLUMN', 'COLUMNS', 'ADDRESS', 'AREAS', 'TRANSPOSE', 'OFFSET', 'HYPERLINK',
  // Financial
  'PMT', 'IPMT', 'PPMT', 'PV', 'FV', 'NPV', 'IRR', 'XIRR', 'XNPV', 'RATE', 'NPER', 'SLN',
  'DB', 'DDB', 'SYD',
  // Information
  'ISBLANK', 'ISERROR', 'ISERR', 'ISNA', 'ISNUMBER', 'ISTEXT', 'ISNONTEXT', 'ISLOGICAL',
  'ISREF', 'ISEVEN', 'ISODD', 'ISFORMULA', 'N', 'NA', 'TYPE', 'CELL',
]);

/**
 * Classifies a single formula string as SAFE (kept as a live formula) or
 * UNSAFE (must be flattened to its static value).
 * @param {string} formula - Formula string as returned by Range.getFormulas(), including the leading '='.
 * @param {Set<string>} excludedSheetNames - Names of sheets that will NOT exist in the exported file.
 * @param {Map<string,string>} namedRangeSheetNames - Named range name -> sheet name its range lives on.
 * @returns {'SAFE'|'UNSAFE'}
 */
function classifyFormula_(formula, excludedSheetNames, namedRangeSheetNames) {
  const stripped = stripStringLiterals_(formula);
  const functionNames = extractFunctionNames_(stripped);

  if (functionNames.some((name) => ALWAYS_FLATTEN_FUNCTIONS.has(name))) return 'UNSAFE';
  if (functionNames.some((name) => UNSAFE_SHEETS_ONLY_FUNCTIONS.has(name))) return 'UNSAFE';
  if (functionNames.some((name) => !KNOWN_EXCEL_COMPATIBLE_FUNCTIONS.has(name))) return 'UNSAFE';

  const referencedSheets = extractSheetQualifiedReferences_(stripped);
  if (referencedSheets.some((name) => excludedSheetNames.has(name))) return 'UNSAFE';

  const bareIdentifiers = extractBareIdentifiers_(stripped);
  for (const id of bareIdentifiers) {
    const sheetName = namedRangeSheetNames.get(id);
    if (sheetName && excludedSheetNames.has(sheetName)) return 'UNSAFE';
  }

  return 'SAFE';
}

/**
 * Determines, for every row of a sheet, which 0-based column indices must be
 * flattened to a static value for the export: either the cell holds its own
 * UNSAFE formula, or it holds no formula of its own but a non-blank value
 * (a spill cell or a plain literal — see the block comment above
 * `flattenUnsafeFormulas_` in SpreadsheetDuplicator.js for why those are
 * indistinguishable and both get rewritten).
 *
 * Formula-less cells are only selected when the sheet contains at least one
 * UNSAFE formula. The reason to rewrite them at all is that flattening an
 * anchor drops its spill, and a spill can never cross sheets — so a sheet
 * with no unsafe formula has no anchor to flatten, nothing to protect, and
 * is left completely untouched (which also avoids needlessly re-parsing
 * every literal in it through `setValues()`).
 *
 * This is the single source of truth for "which cells get their value frozen
 * into the export" — reused by both `flattenUnsafeFormulas_` (which does the
 * freezing) and `buildCalculationWatchLists_` in CalculationWaiter.js (which
 * needs to know, before freezing, whether any of those specific cells are
 * still showing the "Loading..." placeholder).
 * @param {string[][]} formulas - From Range.getFormulas().
 * @param {any[][]} values - The corresponding grid from Range.getValues().
 * @param {Set<string>} excludedSheetNames
 * @param {Map<string,string>} namedRangeSheetNames
 * @returns {number[][]} Per row, the 0-based column indices to flatten, ascending.
 */
function getFlattenColumnsByRow_(formulas, values, excludedSheetNames, namedRangeSheetNames) {
  const unsafeByRow = formulas.map((row) => row.map(
    (formula) => formula !== '' && classifyFormula_(formula, excludedSheetNames, namedRangeSheetNames) === 'UNSAFE'
  ));
  const sheetHasUnsafeFormula = unsafeByRow.some((row) => row.includes(true));

  return formulas.map((formulaRow, r) => {
    /** @type {number[]} */
    const cols = [];
    formulaRow.forEach((formula, c) => {
      const value = values[r][c];
      const isFormulaLessValue = !formula && value !== '' && value !== null;
      if (formula ? unsafeByRow[r][c] : sheetHasUnsafeFormula && isFormulaLessValue) cols.push(c);
    });
    return cols;
  });
}
