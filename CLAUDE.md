# xlsx-exporter

Google Apps Script library that exports a Google Sheets spreadsheet to a real `.xlsx` file, automatically flattening formulas that won't survive the trip to Excel (custom Apps Script functions, `IMPORTRANGE`, `GOOGLEFINANCE`, `QUERY`, dynamic/spill functions, and any formula referencing a sheet excluded from the export) to their last computed static value, while leaving Excel-compatible formulas live. See `README.md` for usage, deployment, and an architecture diagram.

## Commands

Run `npm run check` after every edit — it's the aggregate of the two below, takes seconds, and needs no network.

- `npm run typecheck` — `tsc --noEmit` over `src/` (`jsconfig.json`, `checkJs` + `strict`, real `@types/google-apps-script`). Verifies JSDoc types, nullability (e.g. `getSheetByName()`), and cross-file references in the shared global scope.
- `npm run lint` — ESLint (`eslint.config.mjs`): `no-var`, `prefer-const`, `eqeqeq`, plus `eslint-plugin-jsdoc` requiring a typed JSDoc on every function. `no-undef` is off on purpose — `tsc` owns undefined-name checking.
- Not wired into a pre-commit hook or CI yet.

## Layout

Everything pushed to Apps Script lives in `src/` (`.clasp.json` has `"rootDir": "src"`), so tooling config at the repo root (`jsconfig.json`, `eslint.config.mjs`, `package.json`, docs) is never pushed. `appsscript.json` must stay inside `src/`. Design rationale lives in `docs/decisions/` (ADRs).

## Project-specific conventions

- Plain `.js` Apps Script files — every function has a JSDoc comment with typed `@param`/`@returns`, per the global JS conventions (enforced by `npm run lint`).
- Apps Script concatenates every `.js` file in `src/` into one global scope (no import system). The file split (`Main.js`, `FormulaClassifier.js`, `FormulaParser.js`, `SpreadsheetDuplicator.js`, `SheetUtils.js`, `CalculationWaiter.js`, `XlsxFetch.js`, `DriveUtils.js`) is purely organizational. Top-level `const`s must only be read inside functions (load order isn't guaranteed unless `filePushOrder` is set).
- **Public API is only `exportSpreadsheetToXlsxBlob` and `exportSpreadsheetToXlsxFile`.** Every other helper ends in `_` so Apps Script hides it from library consumers; name new helpers the same way. See [ADR 0006](docs/decisions/0006-private-helpers-trailing-underscore.md).
- `@types/google-apps-script` (plus `typescript`, `eslint`, `@eslint/js`, `eslint-plugin-jsdoc`) are dev-only dependencies (`npm install`) and are never pushed to Apps Script.

## Architecture summary

`Main.js` is the public entry point (`exportSpreadsheetToXlsxBlob`, `exportSpreadsheetToXlsxFile`). It first waits (`CalculationWaiter.js`) for any source cell that would be flattened to finish calculating, then duplicates the source spreadsheet in Drive (`SpreadsheetDuplicator.js`), deletes excluded sheets from the copy, then flattens unsafe cells in the copy by reading formulas/values from the **untouched source** and writing static values into the **duplicate** — never the reverse, and the source is never mutated. Each cell's safe/unsafe classification (`FormulaClassifier.js`) relies on regex-based extraction of function names and sheet references from the formula text (`FormulaParser.js`); `FormulaClassifier.js` also exposes `getFlattenColumnsByRow_`, the single source of truth for "which cells get frozen into the export," shared by both the flatten step and the calculation wait. Once flattening is done, `XlsxFetch.js` fetches the real `.xlsx` bytes via Google's native `/export?format=xlsx` endpoint, and the temporary Drive copy is deleted with retry in a `finally` block. `DriveUtils.js` also sweeps orphaned temp copies from previous runs and saves the resulting Blob to a Drive folder. `SheetUtils.js` holds the small sheet helpers shared across files (`getRequiredSheet_`, `buildNamedRangeSheetMap_`).

## Key non-obvious design decisions (read the linked ADR before changing the related logic)

- **Only cells holding their own SAFE live formula are left untouched; on a sheet with at least one UNSAFE formula, every other non-blank cell is rewritten from the source** (spill cells hold no content of their own and go blank once their anchor is flattened; sheets with no unsafe formula are skipped). → [0001](docs/decisions/0001-rewrite-formula-less-cells-from-source.md)
- **Data validation is cleared on every rewritten range before the value is written**, or programmatic writes can throw. Found via real testing — re-test before removing. → [0002](docs/decisions/0002-clear-data-validation-before-writing.md)
- **`IMPORTRANGE` cells never depend on the duplicate's own evaluation** — values always come from the authorized source. → [0003](docs/decisions/0003-importrange-values-come-from-the-source.md)
- **Temp-copy cleanup can't rely on `finally`** (a hard kill skips it): `cleanUpOrphanedExportTempFiles_` sweeps leftovers on every export, and `makeCopy()` must keep passing an explicit destination folder. Verify the sweep by hand after touching it (see the ADR). → [0004](docs/decisions/0004-orphaned-temp-copy-sweep.md)
- **Calculation wait**: no public status API, so it polls for the `"Loading..."` placeholder and requires two consecutive identical, placeholder-free reads (absence of the placeholder isn't proof of completion — don't revert to a single read); it watches only the cells that will be flattened, and runs before the Drive duplicate. → [0005](docs/decisions/0005-calculation-wait.md)

## Testing

No automated test framework and no test harness in the repo: after `clasp push`, the consuming scratch project uses the library's **Head** version and calls the export directly. See the README's "Testing" section for the recommended scratch-spreadsheet layout that exercises safe formulas, custom functions, `IMPORTRANGE`, excluded-sheet references, and `ARRAYFORMULA` spills.

## Deployment

Script ID is in `.clasp.json`. Push with `clasp push` (only when explicitly asked). Testing uses the library's Head version, so no new deployment is needed until a release is wanted (**Deploy > New deployment**, type **Library**). See the README's "Deploying the library" and "Connecting the library in another project" sections — consuming projects must also declare this library's `oauthScopes` in their own manifest, since Apps Script does not scan a library's code for scope auto-detection.
