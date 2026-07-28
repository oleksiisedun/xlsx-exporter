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
  'SINH', 'COSH', 'TANH',
  // Statistical
  'AVERAGE', 'AVERAGEA', 'AVERAGEIF', 'AVERAGEIFS', 'COUNT', 'COUNTA', 'COUNTBLANK',
  'COUNTIF', 'COUNTIFS', 'MAX', 'MAXA', 'MAXIFS', 'MIN', 'MINA', 'MINIFS', 'MEDIAN',
  'MODE', 'STDEV', 'STDEVA', 'STDEVP', 'STDEVPA', 'VAR', 'VARA', 'VARP', 'VARPA', 'LARGE',
  'SMALL', 'RANK', 'PERCENTILE', 'QUARTILE', 'TRIMMEAN', 'CORREL', 'COVAR', 'FORECAST',
  'TREND', 'GROWTH', 'SLOPE', 'INTERCEPT', 'RSQ', 'NORMDIST', 'NORMSDIST', 'NORMINV',
  'NORMSINV',
  // Logical
  'IF', 'IFS', 'IFERROR', 'IFNA', 'AND', 'OR', 'NOT', 'XOR', 'TRUE', 'FALSE', 'SWITCH',
  // Text
  'CONCATENATE', 'CONCAT', 'TEXTJOIN', 'LEFT', 'RIGHT', 'MID', 'LEN', 'LOWER', 'UPPER',
  'PROPER', 'TRIM', 'SUBSTITUTE', 'REPLACE', 'FIND', 'SEARCH', 'TEXT', 'VALUE', 'REPT',
  'CHAR', 'CODE', 'EXACT', 'CLEAN', 'DOLLAR', 'FIXED', 'UNICHAR', 'UNICODE',
  // Date/Time
  'DATE', 'DATEVALUE', 'TIME', 'TIMEVALUE', 'NOW', 'TODAY', 'YEAR', 'MONTH', 'DAY',
  'HOUR', 'MINUTE', 'SECOND', 'WEEKDAY', 'WEEKNUM', 'DAYS', 'DAYS360', 'NETWORKDAYS',
  'WORKDAY', 'EDATE', 'EOMONTH', 'DATEDIF',
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
function classifyFormula(formula, excludedSheetNames, namedRangeSheetNames) {
  const stripped = stripStringLiterals(formula);
  const functionNames = extractFunctionNames(stripped);

  if (functionNames.some((name) => ALWAYS_FLATTEN_FUNCTIONS.has(name))) return 'UNSAFE';
  if (functionNames.some((name) => UNSAFE_SHEETS_ONLY_FUNCTIONS.has(name))) return 'UNSAFE';
  if (functionNames.some((name) => !KNOWN_EXCEL_COMPATIBLE_FUNCTIONS.has(name))) return 'UNSAFE';

  const referencedSheets = extractSheetQualifiedReferences(stripped);
  if (referencedSheets.some((name) => excludedSheetNames.has(name))) return 'UNSAFE';

  const bareIdentifiers = extractBareIdentifiers(stripped);
  for (const id of bareIdentifiers) {
    const sheetName = namedRangeSheetNames.get(id);
    if (sheetName && excludedSheetNames.has(sheetName)) return 'UNSAFE';
  }

  return 'SAFE';
}

/**
 * Determines which 0-based column indices in a single row must be flattened
 * to a static value for the export: either the cell holds its own UNSAFE
 * formula, or it holds no formula of its own but a non-blank value (a spill
 * cell or a plain literal — see the block comment above
 * `flattenUnsafeFormulas` in SpreadsheetDuplicator.js for why those are
 * indistinguishable and both get rewritten). This is the single source of
 * truth for "which cells get their value frozen into the export" — reused
 * by both `flattenUnsafeFormulas` (which does the freezing) and
 * `buildCalculationWatchLists` in CalculationWaiter.js (which needs to know,
 * before freezing, whether any of those specific cells are still showing
 * the "Loading..." placeholder).
 * @param {string[]} formulaRow - One row from Range.getFormulas().
 * @param {Array} valueRow - The corresponding row from Range.getValues().
 * @param {Set<string>} excludedSheetNames
 * @param {Map<string,string>} namedRangeSheetNames
 * @returns {number[]} 0-based column indices, ascending.
 */
function getFlattenColumnIndices(formulaRow, valueRow, excludedSheetNames, namedRangeSheetNames) {
  const cols = [];
  for (let c = 0; c < formulaRow.length; c++) {
    const formula = formulaRow[c];
    const value = valueRow[c];
    if (!formula) {
      if (value !== '' && value !== null) cols.push(c);
    } else if (classifyFormula(formula, excludedSheetNames, namedRangeSheetNames) === 'UNSAFE') {
      cols.push(c);
    }
  }
  return cols;
}
