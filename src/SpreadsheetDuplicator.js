/**
 * Creates a full copy of the source spreadsheet file in Drive, placed
 * explicitly in the source's own parent folder (or Drive's root, if the
 * source has none). Without an explicit destination, `File.makeCopy(name)`
 * drops the copy in the current user's Drive root regardless of where the
 * source lives — which would silently break `cleanUpOrphanedExportTempFiles_`
 * in DriveUtils.js, since it searches the source's parent folder(s) for
 * leftover copies and would never find any placed in root instead.
 * @param {string} spreadsheetId
 * @param {string} copyName
 * @returns {GoogleAppsScript.Drive.File}
 */
function duplicateSpreadsheetFile_(spreadsheetId, copyName) {
  const sourceFile = DriveApp.getFileById(spreadsheetId);
  const parentIterator = sourceFile.getParents();
  const destinationFolder = parentIterator.hasNext() ? parentIterator.next() : DriveApp.getRootFolder();
  return sourceFile.makeCopy(copyName, destinationFolder);
}

/**
 * Deletes the given sheets (by name) from a spreadsheet, if present.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string[]} sheetNames
 * @returns {void}
 */
function deleteSheetsByName_(spreadsheet, sheetNames) {
  sheetNames.forEach((name) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) spreadsheet.deleteSheet(sheet);
  });
}

/**
 * For every included sheet, reads formulas+values from the SOURCE spreadsheet
 * (read-only — the source is never mutated), classifies each formula, and
 * writes flattened values into the corresponding sheet of the DUPLICATE
 * spreadsheet. Cells holding their own SAFE live formula are left completely
 * untouched; every other non-blank cell of a sheet that contains at least one
 * UNSAFE formula is (re)written from the source value, for two different
 * reasons:
 * - A cell with its own UNSAFE formula is flattened because that formula
 *   won't survive the export.
 * - A cell with NO formula of its own (Range.getFormulas() returns "") is
 *   also (re)written, because it's indistinguishable from a manually-typed
 *   literal without a formula engine: it may equally be the visual result of
 *   an array-producing formula anchored elsewhere (ARRAYFORMULA, QUERY,
 *   SORT, IMPORTRANGE, ...) spilling into it. Such spilled cells hold no
 *   real content of their own — the moment their anchor is flattened from a
 *   formula into a literal, Sheets drops the spill and they go blank. Since
 *   we already have the source's correct value for every such cell in hand,
 *   always re-writing it guarantees no data loss; re-writing an actual
 *   literal with its own unchanged value is a harmless no-op. A sheet with
 *   no UNSAFE formula has no anchor to flatten, so it is skipped entirely
 *   (see `getFlattenColumnsByRow_`).
 *
 * Because the source is read directly (not the duplicate), this stays
 * correct even though excluded sheets have already been deleted from the
 * duplicate at this point: a formula referencing an excluded sheet would
 * show #REF! if read from the duplicate, but here it's read from the source
 * where that sheet still exists and the formula still computes correctly,
 * so the static value written into the duplicate is the correct one.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} duplicateSpreadsheet
 * @param {string[]} includedSheetNames
 * @param {Set<string>} excludedSheetNames
 * @param {Map<string,string>} namedRangeSheetNames
 * @returns {void}
 */
function flattenUnsafeFormulas_(sourceSpreadsheet, duplicateSpreadsheet, includedSheetNames, excludedSheetNames, namedRangeSheetNames) {
  includedSheetNames.forEach((sheetName) => {
    const sourceSheet = getRequiredSheet_(sourceSpreadsheet, sheetName);
    const dataRange = sourceSheet.getDataRange();
    if (dataRange.getNumRows() === 0 || dataRange.getNumColumns() === 0) return;

    const values = dataRange.getValues();
    const flattenColsByRow = getFlattenColumnsByRow_(dataRange.getFormulas(), values, excludedSheetNames, namedRangeSheetNames);
    if (flattenColsByRow.every((cols) => cols.length === 0)) return;

    writeFlattenedRectangles_(getRequiredSheet_(duplicateSpreadsheet, sheetName), mergeColumnsIntoRectangles_(flattenColsByRow), values);
  });
}

/**
 * @typedef {Object} CellRectangle
 * @property {number} row - 0-based top row.
 * @property {number} col - 0-based left column.
 * @property {number} numRows
 * @property {number} numCols
 */

/**
 * Merges per-row column selections into the fewest rectangles: contiguous
 * columns within a row form a run, and runs with identical column extent on
 * consecutive rows are stacked into one rectangle. This keeps the number of
 * Sheets API calls proportional to the number of distinct blocks rather than
 * to rows x runs, which is what keeps large sheets inside Apps Script's
 * 6-minute execution limit.
 * @param {number[][]} flattenColsByRow - Per row, 0-based column indices, ascending.
 * @returns {CellRectangle[]}
 */
function mergeColumnsIntoRectangles_(flattenColsByRow) {
  /** @type {CellRectangle[]} */
  const rectangles = [];
  /** @type {Map<string, CellRectangle>} Rectangles that end on the previous row, keyed by column extent. */
  let openRectangles = new Map();

  flattenColsByRow.forEach((cols, row) => {
    /** @type {Map<string, CellRectangle>} */
    const nextOpenRectangles = new Map();
    let i = 0;
    while (i < cols.length) {
      let j = i;
      while (j + 1 < cols.length && cols[j + 1] === cols[j] + 1) j++;
      const col = cols[i];
      const numCols = cols[j] - col + 1;
      const key = `${col}:${numCols}`;

      let rectangle = openRectangles.get(key);
      if (rectangle) {
        rectangle.numRows++;
      } else {
        rectangle = { row, col, numRows: 1, numCols };
        rectangles.push(rectangle);
      }
      nextOpenRectangles.set(key, rectangle);
      i = j + 1;
    }
    openRectangles = nextOpenRectangles;
  });
  return rectangles;
}

/**
 * Writes static values into each rectangle, clearing data validation on it
 * first. A rule that validates against a list sourced from a now-deleted
 * excluded sheet (or simply doesn't accept the flattened value's type) would
 * otherwise make Range.setValues() throw, since Sheets enforces "reject
 * invalid input" validation on programmatic writes just like on manual entry.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dupSheet
 * @param {CellRectangle[]} rectangles
 * @param {any[][]} values - The source data range's values, indexed from the sheet's top-left cell (A1).
 * @returns {void}
 */
function writeFlattenedRectangles_(dupSheet, rectangles, values) {
  rectangles.forEach(({ row, col, numRows, numCols }) => {
    const range = dupSheet.getRange(row + 1, col + 1, numRows, numCols);
    range.clearDataValidations();
    range.setValues(values.slice(row, row + numRows).map((rowValues) => rowValues.slice(col, col + numCols)));
  });
}
