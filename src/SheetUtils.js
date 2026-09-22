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

/**
 * Converts A1-notation column letters to a 0-based column index
 * ("A" -> 0, "Z" -> 25, "AA" -> 26). Case-insensitive.
 * @param {string} letters
 * @returns {number}
 */
function columnLettersToIndex_(letters) {
  let index = 0;
  for (const ch of letters.toUpperCase()) index = index * 26 + (ch.charCodeAt(0) - 64);
  return index - 1;
}

/**
 * True if the inclusive column range [start, end] overlaps any span in `spans`.
 * @param {{start: number, end: number}[]} spans
 * @param {number} start
 * @param {number} end
 * @returns {boolean}
 */
function spansOverlap_(spans, start, end) {
  return spans.some((span) => start <= span.end && end >= span.start);
}
