# xlsx-exporter

Google Apps Script library that exports a Google Sheets spreadsheet to a real `.xlsx` file, automatically flattening formulas that won't survive the trip to Excel (custom Apps Script functions, `IMPORTRANGE`, `GOOGLEFINANCE`, `QUERY`, dynamic/spill functions, and any formula referencing a sheet excluded from the export) to their last computed static value, while leaving Excel-compatible formulas live. See `README.md` for usage, deployment, and an architecture diagram.

## Project-specific conventions

- Plain `.js` Apps Script files — every function has a JSDoc comment with typed `@param`/`@returns`, per the global JS conventions.
- Apps Script concatenates every `.js` file in the project into one global scope (no import system). The file split (`Main.js`, `FormulaClassifier.js`, `FormulaParser.js`, `SpreadsheetDuplicator.js`, `XlsxExport.js`, `DriveUtils.js`, `Test.js`) is purely organizational.

## Architecture summary

`Main.js` is the public entry point (`exportSpreadsheetToXlsxBlob`, `exportSpreadsheetToXlsxFile`). It duplicates the source spreadsheet in Drive (`SpreadsheetDuplicator.js`), deletes excluded sheets from the copy, then flattens unsafe cells in the copy by reading formulas/values from the **untouched source** and writing static values into the **duplicate** — never the reverse, and the source is never mutated. Each cell's safe/unsafe classification (`FormulaClassifier.js`) relies on regex-based extraction of function names and sheet references from the formula text (`FormulaParser.js`). Once flattening is done, `XlsxExport.js` fetches the real `.xlsx` bytes via Google's native `/export?format=xlsx` endpoint, and the temporary Drive copy is deleted in a `finally` block. `DriveUtils.js` is a small helper to save the resulting Blob to a Drive folder.

## Key non-obvious design decisions (read before changing flatten logic)

- **Any cell with no formula of its own is also rewritten from the source**, not just unsafe-formula cells. This is because `ARRAYFORMULA`/`QUERY`/`SORT`/`IMPORTRANGE`-style formulas store their formula text only in the top-left anchor cell — every other cell they visually fill has an empty formula string and no real stored content. The instant the anchor is flattened to a literal, Sheets drops the spill and those cells go blank. Since Apps Script's API can't report a formula's spill boundaries, the only reliable fix is to always re-write every non-blank, formula-less cell from the source's captured value — a real literal gets rewritten with its own unchanged value (harmless no-op), while a spill artifact is correctly preserved. See the block comment above `flattenUnsafeFormulas` in `SpreadsheetDuplicator.js`.
- **Data validation is cleared on every rewritten cell before the value is written.** Sheets enforces "reject invalid input" validation on programmatic `setValues()` calls, not just manual entry. A validation rule whose source list lives on a sheet we just deleted (or that simply doesn't accept the flattened value's type) would otherwise make the write throw. This was found via real testing, not anticipated up front — don't remove it without re-testing against a sheet that has dropdown validation sourced from an excluded sheet.
- **`IMPORTRANGE` cells never depend on the duplicate's own evaluation.** A Drive copy is a new file ID and starts without `IMPORTRANGE` authorization, but since `IMPORTRANGE` is always classified unsafe, its cells are always flattened using the value read from the already-authorized source — the duplicate's own (unauthorized) evaluation of that formula is never read.
- **Only cells holding their own SAFE live formula are left untouched.** Everything else non-blank gets rewritten, even when that means touching more cells than strictly necessary — correctness (no data loss) is prioritized over minimizing write volume or preserving data-validation fidelity on cells incidentally caught by the blanket rewrite.

## Testing

No automated test framework exists for Apps Script. `Test.js` has `testEndToEndExport()` — fill in a real scratch spreadsheet ID and Drive folder ID, then run it from the Apps Script editor. See the README's "Testing" section for the recommended scratch-spreadsheet layout that exercises safe formulas, custom functions, `IMPORTRANGE`, excluded-sheet references, and `ARRAYFORMULA` spills.

## Deployment

Script ID is in `.clasp.json`. Push with `clasp push`, then cut a new library deployment (**Deploy > New deployment**, type **Library**) from the Apps Script editor. See the README's "Deploying the library" and "Connecting the library in another project" sections — consuming projects must also declare this library's `oauthScopes` in their own manifest, since Apps Script does not scan a library's code for scope auto-detection.
