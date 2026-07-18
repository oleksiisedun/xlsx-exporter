/**
 * Creates a full copy of the source spreadsheet file in Drive.
 * @param {string} spreadsheetId
 * @param {string} copyName
 * @returns {GoogleAppsScript.Drive.File}
 */
function duplicateSpreadsheetFile(spreadsheetId, copyName) {
  return DriveApp.getFileById(spreadsheetId).makeCopy(copyName);
}

/**
 * Deletes the given sheets (by name) from a spreadsheet, if present.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string[]} sheetNames
 * @returns {void}
 */
function deleteSheetsByName(spreadsheet, sheetNames) {
  sheetNames.forEach((name) => {
    const sheet = spreadsheet.getSheetByName(name);
    if (sheet) spreadsheet.deleteSheet(sheet);
  });
}

/**
 * Builds a map from named-range name to the name of the sheet its range lives on.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @returns {Map<string,string>}
 */
function buildNamedRangeSheetMap(spreadsheet) {
  const map = new Map();
  spreadsheet.getNamedRanges().forEach((namedRange) => {
    map.set(namedRange.getName(), namedRange.getRange().getSheet().getName());
  });
  return map;
}

/**
 * For every included sheet, reads formulas+values from the SOURCE spreadsheet
 * (read-only — the source is never mutated), classifies each formula, and
 * writes flattened values into the corresponding sheet of the DUPLICATE
 * spreadsheet. Cells holding their own SAFE live formula are left completely
 * untouched; everything else non-blank is (re)written from the source value,
 * for two different reasons:
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
 *   literal with its own unchanged value is a harmless no-op.
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
function flattenUnsafeFormulas(sourceSpreadsheet, duplicateSpreadsheet, includedSheetNames, excludedSheetNames, namedRangeSheetNames) {
  includedSheetNames.forEach((sheetName) => {
    const sourceSheet = sourceSpreadsheet.getSheetByName(sheetName);
    const dataRange = sourceSheet.getDataRange();
    const numRows = dataRange.getNumRows();
    const numCols = dataRange.getNumColumns();
    if (numRows === 0 || numCols === 0) return;

    const formulas = dataRange.getFormulas();
    const values = dataRange.getValues();
    const dupSheet = duplicateSpreadsheet.getSheetByName(sheetName);

    for (let r = 0; r < numRows; r++) {
      const flattenCols = [];
      for (let c = 0; c < numCols; c++) {
        const formula = formulas[r][c];
        const value = values[r][c];
        if (!formula) {
          if (value !== '' && value !== null) flattenCols.push(c);
        } else if (classifyFormula(formula, excludedSheetNames, namedRangeSheetNames) === 'UNSAFE') {
          flattenCols.push(c);
        }
      }
      writeFlattenedRunsForRow(dupSheet, r, flattenCols, values[r]);
    }
  });
}

/**
 * Writes static values into contiguous runs of columns within a single row,
 * clearing data validation on each run first. A rule that validates against
 * a list sourced from a now-deleted excluded sheet (or simply doesn't accept
 * the flattened value's type) would otherwise make Range.setValues() throw,
 * since Sheets enforces "reject invalid input" validation on programmatic
 * writes just like on manual entry.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} dupSheet
 * @param {number} rowIndex - 0-based row index.
 * @param {number[]} flattenCols - 0-based column indices to flatten, ascending.
 * @param {Array} rowValues - The source row's values (one row from Range.getValues()).
 * @returns {void}
 */
function writeFlattenedRunsForRow(dupSheet, rowIndex, flattenCols, rowValues) {
  let i = 0;
  while (i < flattenCols.length) {
    let j = i;
    while (j + 1 < flattenCols.length && flattenCols[j + 1] === flattenCols[j] + 1) j++;
    const startCol = flattenCols[i];
    const runLength = flattenCols[j] - startCol + 1;
    const range = dupSheet.getRange(rowIndex + 1, startCol + 1, 1, runLength);
    range.clearDataValidations();
    range.setValues([rowValues.slice(startCol, startCol + runLength)]);
    i = j + 1;
  }
}
