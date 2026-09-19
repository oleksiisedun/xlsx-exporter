/**
 * Saves a Blob into a Drive folder, returning the created File.
 * @param {GoogleAppsScript.Base.Blob} blob
 * @param {string} folderId
 * @param {string} [fileName] - Overrides the blob's current name if provided.
 * @returns {GoogleAppsScript.Drive.File}
 */
function saveBlobToDriveFolder(blob, folderId, fileName) {
  const folder = DriveApp.getFolderById(folderId);
  return folder.createFile(fileName ? blob.setName(fileName) : blob);
}

/**
 * Trashes a Drive file, retrying with backoff on transient failures (rate
 * limits, brief API outages) so a temporary export copy isn't left behind
 * just because a single setTrashed() call happened to fail.
 * @param {string} fileId
 * @param {number} [maxAttempts]
 * @returns {void}
 */
function deleteFileWithRetry(fileId, maxAttempts) {
  const attempts = maxAttempts || 3;
  for (let attempt = 1; attempt <= attempts; attempt++) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      return;
    } catch (error) {
      if (attempt === attempts) {
        console.warn(`deleteFileWithRetry: failed to trash file ${fileId} after ${attempts} attempt(s): ${error}`);
        return;
      }
      Utilities.sleep(1000 * attempt);
    }
  }
}

/**
 * Sweeps up leftover temporary export copies from previous runs that were
 * never cleaned up. This is a self-healing safety net for the case a
 * `try/finally` can't cover: if Apps Script hard-kills an execution (e.g. the
 * 6-minute timeout, or a manual stop from the Executions dashboard), the
 * process is torn down immediately and `finally` never runs — there is no
 * error to catch and no cleanup code executes at all. Copies are made
 * alongside the source file (`duplicateSpreadsheetFile` in
 * SpreadsheetDuplicator.js passes an explicit destination folder for exactly
 * this reason), so only the source's parent folders need to be swept.
 * Only copies older than maxAgeMs are removed, well past any real export's
 * runtime, so a temp copy from an export still legitimately in progress is
 * never touched.
 * @param {string} spreadsheetId - Source spreadsheet whose parent folders to sweep.
 * @param {string} tempFilePrefix - Name prefix identifying this library's temp copies.
 * @param {number} [maxAgeMs] - Minimum age before an orphaned copy is trashed.
 * @returns {void}
 */
function cleanUpOrphanedExportTempFiles(spreadsheetId, tempFilePrefix, maxAgeMs) {
  const minAge = maxAgeMs || 15 * 60 * 1000;
  const cutoff = Date.now() - minAge;
  const parentIterator = DriveApp.getFileById(spreadsheetId).getParents();
  const parents = [];
  while (parentIterator.hasNext()) parents.push(parentIterator.next());
  if (parents.length === 0) parents.push(DriveApp.getRootFolder());

  parents.forEach((folder) => {
    const files = folder.getFiles();
    while (files.hasNext()) {
      const file = files.next();
      if (!file.getName().startsWith(tempFilePrefix)) continue;
      if (file.getDateCreated().getTime() > cutoff) continue;
      deleteFileWithRetry(file.getId());
    }
  });
}
