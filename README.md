# xlsx-exporter

A Google Apps Script library that exports a Google Sheets spreadsheet to a real `.xlsx` file, automatically detecting formulas that won't survive the trip to Excel and flattening only those to static values — everything else stays a live formula.

## Why

Google Sheets' own "Download as .xlsx" conversion doesn't reliably handle formulas that have no Excel equivalent: custom Apps Script functions, `IMPORTRANGE`, `GOOGLEFINANCE`, `QUERY`, etc. often come out broken, blank, or as `#NAME?` errors. This library pre-processes those specific cells — replacing them with their last computed value — before generating the export, so the resulting file has no data loss.

It also solves a related problem: if you export only a subset of sheets and a formula on an included sheet references data on an excluded sheet, that formula would otherwise break (`#REF!`) in the export. This library detects that case and flattens it to a static value too.

## Deploying the library

1. Open the project in the Apps Script editor (`clasp open`, or push first with `clasp push` if you've made local changes).
2. **Deploy > New deployment**, select type **Library**, and create the deployment. Note the **Script ID** shown under **Project Settings** (also the `scriptId` in this repo's `.clasp.json`).
3. Each time you want to release a change, cut a new deployment version (or a new deployment) — consuming projects pin to a specific version number, so old versions keep working until the consumer explicitly updates. For testing, a consuming project can instead select the library's **Head** version, which always runs the latest pushed code.

## Connecting the library in another project

1. In the consuming project's Apps Script editor: **Libraries** (left sidebar) > **Add a library**.
2. Paste the library's Script ID, click **Look up**, pick the version to use, and set an identifier (e.g. `XlsxExporter`) — this identifier is the namespace you'll call functions through.
3. Add the [required scopes](#required-scopes-in-the-consuming-project) below to the consuming project's own `appsscript.json` — the library's own manifest scopes are not inherited automatically.

## Usage

```js
// Export everything
const blob = XlsxExporter.exportSpreadsheetToXlsxBlob({ spreadsheetId: '...' });

// spreadsheetId is optional — omit it to export the active spreadsheet
// (only resolvable from a container-bound script or trigger, e.g. a custom menu item)
const blob = XlsxExporter.exportSpreadsheetToXlsxBlob({});

// Export only specific sheets
const blob = XlsxExporter.exportSpreadsheetToXlsxBlob({
  spreadsheetId: '...',
  includeSheets: ['Summary', 'Data'],
});

// Export everything except some sheets, saving directly to Drive
const file = XlsxExporter.exportSpreadsheetToXlsxFile(
  { spreadsheetId: '...', excludeSheets: ['Internal Notes'] },
  'DRIVE_FOLDER_ID'
);

// Remove columns from a sheet (deleted by default; formulas that read them are flattened)
const blob = XlsxExporter.exportSpreadsheetToXlsxBlob({
  spreadsheetId: '...',
  excludeColumns: { 'Data': ['C', 'F:H'], 'Summary': ['B'] },
});

// ...or just hide them (data stays in the file, nothing referencing them breaks)
const blob = XlsxExporter.exportSpreadsheetToXlsxBlob({
  spreadsheetId: '...',
  excludeColumns: { 'Data': ['C'] },
  excludeColumnsMode: 'hide',
});

// Tune (or skip) the wait for in-progress calculations before export
const blob = XlsxExporter.exportSpreadsheetToXlsxBlob({
  spreadsheetId: '...',
  calculationWaitTimeoutMs: 180000,      // wait up to 3 min (default 2 min)
  calculationWaitPollIntervalMs: 5000,   // check every 5s (default 3s)
});

// Override the base file name (the date/time suffix is still appended)
const blob = XlsxExporter.exportSpreadsheetToXlsxBlob({
  spreadsheetId: '...',
  fileName: 'Data Export', // → "Data Export 28.07.2026 15:51"
});
```

`spreadsheetId` is optional — if omitted, the library falls back to `SpreadsheetApp.getActiveSpreadsheet()`, which only resolves when called from a bound script context (a container-bound script or a simple/installable trigger); calling it without `spreadsheetId` from a standalone script or webapp throws.

`includeSheets` and `excludeSheets` are mutually exclusive — pass at most one. Passing neither exports every sheet. `excludeColumns` maps a sheet name (which must be part of the export) to column specs — `'C'` or `'F:H'`, case-insensitive; `excludeColumnsMode` is `'delete'` (default) or `'hide'`. See [Excluding columns](#excluding-columns) below. The exported file name is always the spreadsheet's name (or the `fileName` option, if given) with the current date/time appended in `DD.MM.YYYY HH:MM` format, using the source spreadsheet's own time zone. Pass `calculationWaitTimeoutMs: 0` to skip the calculation wait entirely — see [Waiting for pending calculations](#waiting-for-pending-calculations) below.

### Required scopes in the consuming project

Apps Script's automatic scope detection only scans a project's own code, not the code of libraries it depends on. Any project that adds this library as a dependency must **also** declare these scopes in its own `appsscript.json`, or calls into this library will fail with an authorization error:

```json
"oauthScopes": [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/script.external_request"
]
```

## Formula classification

Every formula in an included sheet is classified as:

- **SAFE** — kept as a live formula in the export. Anything using only standard functions with a direct Excel equivalent (`SUM`, `VLOOKUP`, `IF`, ...) and not referencing an excluded sheet.
- **UNSAFE** — flattened to its last computed static value. This covers:
  - Google-Sheets-only functions with no Excel equivalent (`IMPORTRANGE`, `GOOGLEFINANCE`, `GOOGLETRANSLATE`, `QUERY`, `IMPORTHTML`/`IMPORTXML`/`IMPORTDATA`/`IMPORTFEED`, `SPARKLINE`, ...).
  - Dynamic/spill-producing functions (`INDIRECT`, `SORT`, `UNIQUE`, `FILTER`, `SEQUENCE`, `RANDARRAY`, `ARRAYFORMULA`) — always flattened, since their target/output shape can't be statically verified.
  - Any function name not recognized as a standard Excel-compatible function — this is what catches custom Apps Script functions without needing to enumerate them.
  - Any formula referencing (directly or via a named range) a sheet that's excluded from the export.
  - Any formula referencing (directly or via a named range) a column deleted via `excludeColumns`.

The known-safe function list and the two blocklists live in `src/FormulaClassifier.js` as the single source of truth — extend them there if testing turns up a false positive.

**Known limitations**:

- A sheet name embedded inside a text argument (e.g. a `QUERY` criteria string) isn't detected by the regex-based reference scanner — but `QUERY` itself is always unsafe, so this doesn't cause data loss in practice.
- Only cell contents are handled. Charts, pivot tables, and conditional-formatting or data-validation rules whose source range lives on an excluded sheet are **not** detected or rewritten, and are likely to break in the export (the sheet is deleted from the temporary copy). Exclude sheets only if nothing else you care about depends on them.

## Excluding columns

`excludeColumns` removes columns from sheets that are otherwise exported. Pick the mode by what you need:

- **`'delete'`** (default) — the columns are gone from the file. Any formula, on any included sheet, that references a deleted column (`C2`, `$C$2`, `A1:C5`, `C:C`, `Data!C2`, a whole-row `2:2`, or a named range touching one) is classified **unsafe** and flattened to its value, because Sheets would either show `#REF!` or silently shrink a range that merely overlaps a deleted column and change its result. Formulas that only touch other columns stay live; Sheets adjusts their references for the shift.
- **`'hide'`** — the columns are only hidden. The data is still in the file (anyone can unhide it), so don't use this to keep data private; use it when charts, pivots or validation depend on the columns, since nothing is deleted and nothing needs flattening.

Columns are removed from the temporary copy after flattening. A sheet can't have every column deleted. Charts, pivot tables and validation/conditional-format rules that reference a deleted column are **not** detected (same limitation as excluded sheets). Design notes: [ADR 0007](docs/decisions/0007-exclude-columns-delete-or-hide.md).

## How export actually happens

The source spreadsheet is **never mutated**. The library duplicates it in Drive, deletes the excluded sheets from the copy, flattens unsafe-formula cells in the copy (writing values read from the untouched source, not the copy — this is what keeps a formula referencing a just-deleted sheet correct instead of showing `#REF!`), deletes or hides any `excludeColumns` columns (last, so flattening still sees original column positions), fetches the real `.xlsx` bytes via Google's native `/export?format=xlsx` endpoint, then deletes the temporary copy (retrying on transient failures).

Because Apps Script can hard-kill an execution (the 6-minute timeout, or a manual stop from the Executions dashboard) without ever running its `finally` block, a temp copy can occasionally be left behind with no code able to clean it up. Every export call opportunistically sweeps the source file's parent folders (via a server-side Drive search) for its own leftover `__xlsx_export_tmp__`-prefixed copies older than 15 minutes and trashes them, so orphans from a previous killed run get cleaned up on the next export rather than accumulating indefinitely.

Only cells holding their own **safe** live formula are left untouched. On any sheet that contains at least one unsafe formula, every other non-blank cell is rewritten from the source value — this includes the unsafe-formula cells themselves, but also any cell with no formula of its own, because that's indistinguishable from a manually-typed literal without a formula engine: it could just as easily be the visual result of an array-producing formula anchored elsewhere (`ARRAYFORMULA`, `QUERY`, `SORT`, `IMPORTRANGE`, ...) spilling into it. Such spilled cells hold no real content — the moment their anchor is flattened from a formula into a literal, Sheets drops the spill and they go blank. Always rewriting them from the source's already-captured value prevents that data loss; rewriting an actual literal with its own unchanged value is a harmless no-op. A sheet with no unsafe formula has no anchor to flatten (a spill can never cross sheets), so it is skipped entirely. Writes are batched into merged rectangles so large sheets stay within Apps Script's 6-minute execution limit.

Each rewritten cell also has its data validation rule cleared before the value is written: Sheets enforces "reject invalid input" validation on programmatic writes too, and a rule sourced from a range on a just-deleted excluded sheet (or one that simply doesn't accept the flattened value's type) would otherwise make the write throw.

```mermaid
graph TD
  Caller["Consuming Apps Script project"] --> Main["Main.js\nexportSpreadsheetToXlsxBlob()"]

  subgraph Lib["xlsx-exporter library (src/)"]
    Main --> Resolve["Main.js\nresolveIncludedSheetNames_()"]
    Main --> Wait["CalculationWaiter.js\nwaitForCalculationsToFinish_()"]
    Main --> Sweep["DriveUtils.js\ncleanUpOrphanedExportTempFiles_()"]
    Main --> Columns["ColumnExclusion.js\nresolveExcludedColumnSpans_()\napplyExcludedColumns_()"]
    Main --> Dup["SpreadsheetDuplicator.js\nduplicateSpreadsheetFile_()"]
    Main --> Flatten["SpreadsheetDuplicator.js\nflattenUnsafeFormulas_()"]
    Wait --> Select
    Flatten --> Select["FormulaClassifier.js\ngetFlattenColumnsByRow_()"]
    Select --> Classify["FormulaClassifier.js\nclassifyFormula_()"]
    Classify --> Parse["FormulaParser.js\nfunction / sheet / column-ref extraction"]
    Main --> Fetch["XlsxFetch.js\nfetchXlsxBlob_()"]
    Main --> Cleanup["DriveUtils.js\ndeleteFileWithRetry_() in finally"]
    Main --> Save["DriveUtils.js\nsaveBlobToDriveFolder_()"]
  end

  Sweep -->|"trashes stale copies"| DriveFolder[("Source's Drive folder")]
  Dup --> DriveCopy[("Temporary Drive copy")]
  Flatten --> DriveCopy
  Columns --> DriveCopy
  Cleanup -->|"trashes"| DriveCopy
  Fetch --> ExportEndpoint[("docs.google.com/.../export?format=xlsx")]
  DriveCopy --> ExportEndpoint
  Fetch --> Main
  Main -->|"Blob"| Caller
  Save --> DriveFolder
```

Because `IMPORTRANGE` cells are always classified unsafe, the library never depends on the temporary Drive copy's own (unauthorized — a Drive copy is a new file ID, so it starts without `IMPORTRANGE` access grants) evaluation of that formula; the static value written into the copy always comes from the already-authorized source spreadsheet.

## Waiting for pending calculations

Custom Apps Script functions and `IMPORTRANGE` compute asynchronously — while a value is still being computed, Sheets shows the placeholder text `"Loading..."` in that cell. If export ran at that exact moment, it would flatten that placeholder into the export as a static value instead of the real result.

Before duplicating the spreadsheet, the library polls the source for exactly the cells that would be flattened (unsafe-formula cells and formula-less non-blank cells — not every cell, since a still-calculating **safe** live formula stays a live formula in the export and is never read). A poll only counts as settled once two consecutive reads, spaced `calculationWaitPollIntervalMs` apart, both show no `"Loading..."` **and** report identical values for every watched cell — the absence of the placeholder on a single read isn't proof a cell is done, since a cross-spreadsheet `IMPORTRANGE` (especially a wide import, or one wrapped in `LET`/`FILTER`/`CHOOSECOLS`) can briefly report a plausible-but-stale value like `0` instead of the placeholder while it's still settling. The library waits until two such reads agree, or until `calculationWaitTimeoutMs` (default 2 minutes) elapses, in which case it throws an error naming the exact sheet/cell still stuck (or noting that values kept changing between polls). Pass `calculationWaitTimeoutMs: 0` to skip this check entirely.

**Known limitation**: a formula-less cell where someone manually typed the literal text `"Loading..."` is indistinguishable from a genuinely pending cell — the export will wait on it and eventually throw a timeout error naming that cell, rather than silently exporting the wrong value.

## Testing

Pure logic has unit tests: run `npm test` (Node's built-in runner; see `test/`). Code that calls `SpreadsheetApp`/`DriveApp` can't run outside Apps Script, so test that end to end against the library's **Head** version: `clasp push`, then in a scratch consuming project (**Libraries** > this library > version **Head**) call `exportSpreadsheetToXlsxBlob` / `exportSpreadsheetToXlsxFile` directly against a scratch spreadsheet and open the result. No new deployment is needed for this.

After changing the orphan sweep (`cleanUpOrphanedExportTempFiles_`), also check it by hand, since a Drive search that silently matches nothing would leave orphans piling up unnoticed: copy the scratch spreadsheet into the same folder with a name starting `__xlsx_export_tmp__`, wait 15+ minutes, run any export, and confirm the copy was trashed.

Suggested scratch spreadsheet layout for a thorough check: a `Data` sheet with safe formulas, a `Custom` sheet with a real custom Apps Script function, an `External` sheet with `IMPORTRANGE`, a `Summary` sheet with a formula referencing `Data!`, a formula/named range referencing a sheet you'll exclude, and an `ARRAYFORMULA` spilling across multiple rows/columns, and an `Excluded` sheet. Confirm: the excluded sheet is absent from the export; safe formulas are still live; `IMPORTRANGE`/custom-function/excluded-referencing cells and the full extent of the `ARRAYFORMULA`'s spilled output are static values, not blank or errored; and the original spreadsheet is completely unchanged afterward. For `excludeColumns`, add a sheet with formulas reading a to-be-deleted column directly, through a range that only overlaps it, via `Sheet!` from another sheet, and through a named range, plus one that reads only surviving columns; confirm the former are static values and the latter stay live and correct after the shift, in both `'delete'` and `'hide'` mode. Also try frozen columns, merged cells and a filter on that sheet. Also worth adding a sheet with **no** unsafe formulas but a few tricky literals (`'00123`, `'=1+1`, a cell with rich text/links) to confirm it comes through untouched.

To verify the calculation wait, add a deliberately slow custom function **defined in the scratch spreadsheet's own bound Apps Script project** (not this library — custom functions execute in the calling spreadsheet's context), e.g.:

```js
function SLOW_VALUE(seconds) {
  Utilities.sleep((seconds || 5) * 1000);
  return 'done';
}
```

Use it in a cell, trigger an edit so it starts recalculating, and immediately run the export. With a generous `calculationWaitTimeoutMs`, confirm the export waits and the real value (`'done'`, not `"Loading..."`) lands in the output. With a very small `calculationWaitTimeoutMs`, confirm the export throws, naming the correct sheet and cell. Also worth checking once: log the cell's raw value while it's still calculating (`console.log(JSON.stringify(value))`) to confirm the exact placeholder text matches `CALCULATION_LOADING_PLACEHOLDER` in `CalculationWaiter.js` for your environment/locale.

## Development

Everything pushed to Apps Script lives in `src/` (`.clasp.json` sets `"rootDir": "src"`); tooling config and docs at the repo root are never pushed. After `npm install`, run `npm run check` (type-check with `tsc` + ESLint) after every edit — see [CLAUDE.md](CLAUDE.md) for what each script enforces. Only `exportSpreadsheetToXlsxBlob` and `exportSpreadsheetToXlsxFile` are public; every other helper ends in `_` so Apps Script keeps it private to the library. Design rationale for the non-obvious parts lives in [`docs/decisions/`](docs/decisions/).
