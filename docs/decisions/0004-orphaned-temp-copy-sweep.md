# 0004 — Orphaned temp-copy sweep (not just `finally`)

## Context

If Apps Script hard-kills an execution (the 6-minute timeout, or a manual stop from the Executions dashboard), the process is torn down immediately and `finally` never runs — no error to catch, no cleanup code executes, and the temp Drive copy is silently orphaned. `deleteFileWithRetry_` (retry with backoff for transient Drive failures) only covers the case where `finally` *does* run.

## Decision

- `cleanUpOrphanedExportTempFiles_` in `src/DriveUtils.js` runs at the start of every export. It sweeps the source file's parent folders for `__xlsx_export_tmp__`-prefixed copies older than 15 minutes (comfortably past any real export's runtime, so a concurrent export's own temp copy is never touched) and trashes them. Don't replace this with "just fix the finally block" — no code inside a killed execution can run.
- `duplicateSpreadsheetFile_` passes an explicit destination folder to `makeCopy()`. Bare `File.makeCopy(name)` places the copy in the user's Drive root regardless of where the source lives; if that regressed, the sweep would look in the wrong folder and silently find nothing.
- Candidates are found with a server-side `searchFiles("title contains ...")` query instead of iterating every file in the folder (O(folder size) on every export). Because a query that silently matches nothing is exactly the failure mode above, the `startsWith` check stays as a safety filter. There is no automated test, so verify by hand after touching this code: make a copy of a scratch spreadsheet in the same folder named with the `__xlsx_export_tmp__` prefix, wait 15+ minutes, run any export, and confirm the copy was trashed.

## Consequences

Orphans from killed runs are cleaned on the next export rather than accumulating.
