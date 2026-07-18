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
