/**
 * Removes double-quoted string literals from a formula so their contents
 * aren't mistaken for function names or sheet references.
 * @param {string} formula
 * @returns {string}
 */
function stripStringLiterals(formula) {
  return formula.replace(/"(?:[^"\\]|\\.)*"/g, '""');
}

/**
 * Extracts the set of function names called directly in a formula.
 * @param {string} formulaWithoutStrings
 * @returns {string[]} Upper-cased, de-duplicated function names.
 */
function extractFunctionNames(formulaWithoutStrings) {
  const names = new Set();
  const re = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g;
  let m;
  while ((m = re.exec(formulaWithoutStrings))) names.add(m[1].toUpperCase());
  return [...names];
}

/** Matches SheetName! or 'Sheet Name'! (with '' as the escaped-quote sequence). */
const SHEET_REF_RE = /'((?:[^']|'')*)'!|([A-Za-z_][A-Za-z0-9_.]*)!/g;

/**
 * Extracts sheet names referenced via SheetName!A1 / 'Sheet Name'!A1:B2 syntax.
 * @param {string} formulaWithoutStrings
 * @returns {string[]} Unescaped sheet names (de-duplicated).
 */
function extractSheetQualifiedReferences(formulaWithoutStrings) {
  const names = new Set();
  let m;
  SHEET_REF_RE.lastIndex = 0;
  while ((m = SHEET_REF_RE.exec(formulaWithoutStrings))) {
    const raw = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
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
function extractBareIdentifiers(formulaWithoutStrings) {
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
