/**
 * Manual test harness — run this function directly from the Apps Script
 * editor. There is no automated test framework available in Apps Script.
 *
 * Fill in the IDs below with a scratch spreadsheet built per the README's
 * testing checklist (Data / Custom / External / Summary / Excluded sheets)
 * and a destination Drive folder before running.
 * @returns {void}
 */
function testEndToEndExport() {
  const spreadsheetId = 'FILL_IN_SCRATCH_SPREADSHEET_ID';
  const folderId = 'FILL_IN_DESTINATION_FOLDER_ID';

  const file = exportSpreadsheetToXlsxFile({ spreadsheetId, excludeSheets: ['Excluded'] }, folderId);
  console.log(`Exported: ${file.getUrl()}`);
}
