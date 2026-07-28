/**
 * Name prefix for temporary Drive copies made during export. Shared between
 * the copy step below and DriveUtils.js's orphan sweep, which identifies
 * leftover copies from previous runs by this prefix.
 * @type {string}
 */
const XLSX_EXPORT_TEMP_FILE_PREFIX = '__xlsx_export_tmp__';

/**
 * @typedef {Object} ExportXlsxOptions
 * @property {string} spreadsheetId - ID of the source Google Sheets spreadsheet. Never mutated.
 * @property {string[]} [includeSheets] - Sheet names to include. Mutually exclusive with excludeSheets.
 * @property {string[]} [excludeSheets] - Sheet names to exclude. Mutually exclusive with includeSheets.
 * @property {string} [fileName] - Base file name (no extension) for the export; defaults to the source spreadsheet's name. The current date/time is always appended.
 */

/**
 * Exports a Google Sheets spreadsheet to a real .xlsx Blob. Formulas that use a
 * Google-Sheets-only function, a dynamic/spill-producing function (INDIRECT,
 * SORT, UNIQUE, ...), or that reference a sheet excluded from the export are
 * flattened to their last computed static value; all other formulas remain
 * live in the export. The source spreadsheet is only ever read, never mutated
 * — all writes happen on a temporary Drive copy, which is deleted afterward
 * even if an error occurs.
 * @param {ExportXlsxOptions} options
 * @returns {GoogleAppsScript.Base.Blob}
 */
function exportSpreadsheetToXlsxBlob(options) {
  const { spreadsheetId, includeSheets, excludeSheets, fileName } = options || {};
  if (!spreadsheetId) throw new Error('exportSpreadsheetToXlsxBlob: options.spreadsheetId is required.');

  const sourceSs = SpreadsheetApp.openById(spreadsheetId);
  const allSheetNames = sourceSs.getSheets().map((s) => s.getName());
  const includedSheetNames = resolveIncludedSheetNames(allSheetNames, includeSheets, excludeSheets);
  const excludedSheetNamesSet = new Set(allSheetNames.filter((n) => !includedSheetNames.includes(n)));
  const namedRangeSheetNames = buildNamedRangeSheetMap(sourceSs);

  const baseFileName = fileName || sourceSs.getName();
  const timestampedFileName = buildTimestampedFileName(baseFileName, sourceSs.getSpreadsheetTimeZone());

  cleanUpOrphanedExportTempFiles(spreadsheetId, XLSX_EXPORT_TEMP_FILE_PREFIX);
  const copiedFile = duplicateSpreadsheetFile(spreadsheetId, `${XLSX_EXPORT_TEMP_FILE_PREFIX}${baseFileName}__${Date.now()}`);

  try {
    const dupSs = SpreadsheetApp.openById(copiedFile.getId());
    deleteSheetsByName(dupSs, [...excludedSheetNamesSet]);
    flattenUnsafeFormulas(sourceSs, dupSs, includedSheetNames, excludedSheetNamesSet, namedRangeSheetNames);
    SpreadsheetApp.flush();
    return fetchXlsxBlob(copiedFile.getId()).setName(`${timestampedFileName}.xlsx`);
  } finally {
    deleteFileWithRetry(copiedFile.getId());
  }
}

/**
 * Convenience helper: exports a spreadsheet to xlsx and saves it into a Drive folder.
 * @param {ExportXlsxOptions} options
 * @param {string} folderId - Destination Drive folder ID.
 * @param {string} [fileName] - Overrides the exported file's name (including the date/time suffix that would otherwise be appended).
 * @returns {GoogleAppsScript.Drive.File}
 */
function exportSpreadsheetToXlsxFile(options, folderId, fileName) {
  const blob = exportSpreadsheetToXlsxBlob(options);
  return saveBlobToDriveFolder(blob, folderId, fileName);
}

/**
 * Appends the current date/time to a base file name, in DD.MM.YYYY HH:MM format.
 * @param {string} baseFileName
 * @param {string} timeZone - IANA time zone, e.g. from Spreadsheet.getSpreadsheetTimeZone().
 * @returns {string}
 */
function buildTimestampedFileName(baseFileName, timeZone) {
  const timestamp = Utilities.formatDate(new Date(), timeZone, 'dd.MM.yyyy HH:mm');
  return `${baseFileName} ${timestamp}`;
}

/**
 * Resolves the final ordered list of sheet names to include, preserving the
 * source spreadsheet's original tab order. If neither includeSheets nor
 * excludeSheets is supplied, all sheets are included.
 * @param {string[]} allSheetNames
 * @param {string[]|undefined} includeSheets
 * @param {string[]|undefined} excludeSheets
 * @returns {string[]}
 */
function resolveIncludedSheetNames(allSheetNames, includeSheets, excludeSheets) {
  const hasInclude = Array.isArray(includeSheets) && includeSheets.length > 0;
  const hasExclude = Array.isArray(excludeSheets) && excludeSheets.length > 0;
  if (hasInclude && hasExclude) {
    throw new Error('exportSpreadsheetToXlsxBlob: pass either includeSheets or excludeSheets, not both.');
  }

  const allSet = new Set(allSheetNames);
  if (hasInclude) {
    assertSheetNamesExist(includeSheets, allSet, 'includeSheets');
    return allSheetNames.filter((name) => includeSheets.includes(name));
  }

  if (hasExclude) assertSheetNamesExist(excludeSheets, allSet, 'excludeSheets');
  const excludeSet = new Set(excludeSheets || []);
  const included = allSheetNames.filter((name) => !excludeSet.has(name));
  if (included.length === 0) {
    throw new Error('exportSpreadsheetToXlsxBlob: resulting included-sheet set is empty.');
  }
  return included;
}

/**
 * @param {string[]} names
 * @param {Set<string>} allSheetNamesSet
 * @param {string} optionLabel
 * @returns {void}
 */
function assertSheetNamesExist(names, allSheetNamesSet, optionLabel) {
  const missing = names.filter((name) => !allSheetNamesSet.has(name));
  if (missing.length > 0) {
    throw new Error(`exportSpreadsheetToXlsxBlob: ${optionLabel} contains unknown sheet name(s): ${missing.join(', ')}`);
  }
}
