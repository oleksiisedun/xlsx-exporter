/**
 * @typedef {Object} ColumnSpan
 * @property {number} start - 0-based first column (inclusive).
 * @property {number} end - 0-based last column (inclusive).
 */

/**
 * The columns that will be physically removed from the export, plus the
 * named ranges that point at them. Both are empty unless columns are being
 * deleted (hidden columns keep their data, so nothing referencing them breaks).
 * @typedef {Object} DeletedColumns
 * @property {Map<string, ColumnSpan[]>} spansBySheet - Sheet name -> sorted, non-overlapping spans.
 * @property {Set<string>} affectedNamedRanges - Names of named ranges overlapping a deleted column.
 */

/** @typedef {'delete'|'hide'} ExcludeColumnsMode */

/** Matches one column spec: "C", "$C" or "F:H". */
const COLUMN_SPEC_RE = /^\$?([A-Za-z]{1,3})(?::\$?([A-Za-z]{1,3}))?$/;

/**
 * Validates the `excludeColumnsMode` option, defaulting to 'delete'.
 * @param {string|undefined} mode
 * @returns {ExcludeColumnsMode}
 */
function resolveExcludeColumnsMode_(mode) {
  if (mode === undefined || mode === 'delete') return 'delete';
  if (mode === 'hide') return 'hide';
  throw new Error(`exportSpreadsheetToXlsxBlob: excludeColumnsMode must be 'delete' or 'hide', got "${mode}".`);
}

/**
 * Parses one column spec ("C" or "F:H", case-insensitive) into a 0-based span.
 * @param {string} spec
 * @param {string} sheetName - For error messages.
 * @returns {ColumnSpan}
 */
function parseColumnSpec_(spec, sheetName) {
  const m = typeof spec === 'string' ? COLUMN_SPEC_RE.exec(spec.trim()) : null;
  if (!m) {
    throw new Error(`exportSpreadsheetToXlsxBlob: excludeColumns["${sheetName}"] contains invalid column "${spec}"; use letters like "C" or "F:H".`);
  }
  const start = columnLettersToIndex_(m[1]);
  const end = m[2] ? columnLettersToIndex_(m[2]) : start;
  return { start: Math.min(start, end), end: Math.max(start, end) };
}

/**
 * Sorts spans and merges any that overlap or touch.
 * @param {ColumnSpan[]} spans
 * @returns {ColumnSpan[]}
 */
function mergeColumnSpans_(spans) {
  /** @type {ColumnSpan[]} */
  const merged = [];
  [...spans].sort((a, b) => a.start - b.start).forEach((span) => {
    const last = merged[merged.length - 1];
    if (last && span.start <= last.end + 1) last.end = Math.max(last.end, span.end);
    else merged.push({ ...span });
  });
  return merged;
}

/**
 * Validates the `excludeColumns` option against the source spreadsheet and
 * resolves it to merged 0-based spans per sheet, clipped to each sheet's
 * grid. Fails before any Drive work if a sheet is unknown or not part of the
 * export, a column spec is malformed, or (when deleting) every column of a
 * sheet would be removed, which Sheets refuses to do.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet
 * @param {string[]} includedSheetNames
 * @param {Object<string, string[]>|undefined} excludeColumns - Sheet name -> column specs.
 * @param {ExcludeColumnsMode} mode
 * @returns {Map<string, ColumnSpan[]>}
 */
function resolveExcludedColumnSpans_(sourceSpreadsheet, includedSheetNames, excludeColumns, mode) {
  /** @type {Map<string, ColumnSpan[]>} */
  const spansBySheet = new Map();
  Object.entries(excludeColumns ?? {}).forEach(([sheetName, specs]) => {
    if (!includedSheetNames.includes(sheetName)) {
      throw new Error(`exportSpreadsheetToXlsxBlob: excludeColumns names sheet "${sheetName}", which is not part of the export.`);
    }
    if (!Array.isArray(specs)) {
      throw new Error(`exportSpreadsheetToXlsxBlob: excludeColumns["${sheetName}"] must be an array of column specs.`);
    }
    const maxColumns = getRequiredSheet_(sourceSpreadsheet, sheetName).getMaxColumns();
    const spans = mergeColumnSpans_(specs.map((spec) => parseColumnSpec_(spec, sheetName)))
      .filter((span) => span.start < maxColumns)
      .map((span) => ({ start: span.start, end: Math.min(span.end, maxColumns - 1) }));
    if (spans.length === 0) return;
    if (mode === 'delete' && spans.length === 1 && spans[0].start === 0 && spans[0].end === maxColumns - 1) {
      throw new Error(`exportSpreadsheetToXlsxBlob: excludeColumns would delete every column of "${sheetName}"; Sheets can't leave a sheet with no columns.`);
    }
    spansBySheet.set(sheetName, spans);
  });
  return spansBySheet;
}

/**
 * Builds the classifier's view of what will be deleted. Empty unless `mode`
 * is 'delete'.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet - The source spreadsheet, to look up named ranges.
 * @param {Map<string, ColumnSpan[]>} excludedColumnSpans
 * @param {ExcludeColumnsMode} mode
 * @returns {DeletedColumns}
 */
function buildDeletedColumns_(spreadsheet, excludedColumnSpans, mode) {
  /** @type {DeletedColumns} */
  const deletedColumns = { spansBySheet: new Map(), affectedNamedRanges: new Set() };
  if (mode !== 'delete' || excludedColumnSpans.size === 0) return deletedColumns;

  deletedColumns.spansBySheet = excludedColumnSpans;
  spreadsheet.getNamedRanges().forEach((namedRange) => {
    const range = namedRange.getRange();
    const spans = excludedColumnSpans.get(range.getSheet().getName());
    const start = range.getColumn() - 1;
    const end = start + range.getNumColumns() - 1;
    if (spans && spansOverlap_(spans, start, end)) {
      deletedColumns.affectedNamedRanges.add(namedRange.getName());
    }
  });
  return deletedColumns;
}

/**
 * Deletes or hides the excluded columns on the DUPLICATE spreadsheet. Must
 * run after `flattenUnsafeFormulas_`, which addresses cells by their original
 * column positions. Spans are deleted right to left so earlier indices stay
 * valid.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} duplicateSpreadsheet
 * @param {Map<string, ColumnSpan[]>} excludedColumnSpans
 * @param {ExcludeColumnsMode} mode
 * @returns {void}
 */
function applyExcludedColumns_(duplicateSpreadsheet, excludedColumnSpans, mode) {
  excludedColumnSpans.forEach((spans, sheetName) => {
    const sheet = getRequiredSheet_(duplicateSpreadsheet, sheetName);
    if (mode === 'hide') {
      spans.forEach(({ start, end }) => sheet.hideColumns(start + 1, end - start + 1));
    } else {
      [...spans].reverse().forEach(({ start, end }) => sheet.deleteColumns(start + 1, end - start + 1));
    }
  });
}
