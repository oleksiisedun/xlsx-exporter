/**
 * Fetches the .xlsx bytes for a Google Sheets file via the native export endpoint.
 * @param {string} spreadsheetFileId
 * @returns {GoogleAppsScript.Base.Blob}
 */
function fetchXlsxBlob_(spreadsheetFileId) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetFileId}/export?format=xlsx`;
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
    muteHttpExceptions: true,
  });
  const code = response.getResponseCode();
  if (code !== 200) {
    throw new Error(`fetchXlsxBlob: export request failed with status ${code}: ${response.getContentText().slice(0, 300)}`);
  }
  return response.getBlob();
}
