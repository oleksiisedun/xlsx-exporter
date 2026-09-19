/**
 * Returns the sheet with the given name, throwing a descriptive error instead
 * of returning null (which `Spreadsheet.getSheetByName` does when it's missing).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string} sheetName
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getRequiredSheet_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error(`Sheet "${sheetName}" not found in spreadsheet "${spreadsheet.getName()}".`);
  return sheet;
}

/**
 * Builds a map from named-range name to the name of the sheet its range lives on.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @returns {Map<string,string>}
 */
function buildNamedRangeSheetMap_(spreadsheet) {
  const map = new Map();
  spreadsheet.getNamedRanges().forEach((namedRange) => {
    map.set(namedRange.getName(), namedRange.getRange().getSheet().getName());
  });
  return map;
}
