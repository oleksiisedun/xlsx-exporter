/**
 * Literal placeholder text Google Sheets shows (and that SpreadsheetApp's
 * read APIs return as the cell's actual value) while a cell is still being
 * computed asynchronously — e.g. a custom Apps Script function that hasn't
 * finished executing yet, or a cross-spreadsheet IMPORTRANGE still resolving.
 * There is no public Apps Script API that reports spreadsheet calculation
 * status; this string is the only observable signal available to a script.
 * Verify this exact string (ellipsis character vs. three periods, locale)
 * against a real slow custom function before relying on it — see the
 * "Testing" section of README.md.
 * @type {string}
 */
const CALCULATION_LOADING_PLACEHOLDER = 'Loading...';

/**
 * Blocks until every cell that `flattenUnsafeFormulas` (SpreadsheetDuplicator.js)
 * would bake into the export as a static value — i.e. every cell
 * `getFlattenColumnIndices` (FormulaClassifier.js) selects — has finished
 * calculating on the SOURCE spreadsheet, or throws once `timeoutMs` elapses.
 *
 * Only that exact cell set is checked, not every cell in the included
 * sheets: it's the single source of truth for which cells actually get
 * their live formula/value read and frozen into the export, so it's also
 * the only set where a lingering "Loading..." placeholder could get baked
 * in as a static literal. A cell holding a SAFE live formula that's still
 * calculating is irrelevant here — it stays a live formula in the export
 * (never read/frozen), so its transient value doesn't matter.
 *
 * Formulas are read from the source exactly once, up front: unlike values,
 * formula text can't change purely as a result of waiting for calculation
 * to settle, so only values are re-read on each poll.
 *
 * Pass `timeoutMs: 0` to opt out entirely — returns immediately, no reads.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet
 * @param {string[]} includedSheetNames
 * @param {Set<string>} excludedSheetNames
 * @param {Map<string,string>} namedRangeSheetNames
 * @param {number} [timeoutMs] - Max total wait before throwing. Defaults to 120000 (2 min). Uses `??`, not `||`, so 0 is a meaningful opt-out.
 * @param {number} [pollIntervalMs] - Delay between re-checks. Defaults to 3000 (3s); floored at 250ms.
 * @returns {void}
 */
function waitForCalculationsToFinish(sourceSpreadsheet, includedSheetNames, excludedSheetNames, namedRangeSheetNames, timeoutMs, pollIntervalMs) {
  const timeout = timeoutMs ?? 120000;
  if (timeout === 0) return;
  const interval = Math.max(pollIntervalMs ?? 3000, 250);

  const watchLists = buildCalculationWatchLists(sourceSpreadsheet, includedSheetNames, excludedSheetNames, namedRangeSheetNames);
  const deadline = Date.now() + timeout;

  while (true) {
    const stillCalculating = findCellStillCalculating(sourceSpreadsheet, watchLists);
    if (!stillCalculating) return;
    if (Date.now() >= deadline) {
      throw new Error(
        `waitForCalculationsToFinish: "${stillCalculating.sheetName}"!${stillCalculating.a1Notation} still shows ` +
        `"${CALCULATION_LOADING_PLACEHOLDER}" after waiting ${timeout}ms; aborting export before it bakes that ` +
        `placeholder into a static value. Pass a larger calculationWaitTimeoutMs if this cell's source (custom ` +
        `function, IMPORTRANGE, ...) genuinely needs longer, or calculationWaitTimeoutMs: 0 to skip this check.`
      );
    }
    Utilities.sleep(interval);
  }
}

/**
 * Precomputes, once per included sheet, the exact 0-based [row, col] pairs
 * that `getFlattenColumnIndices` says will be baked into the export as a
 * static value. Built once because the underlying formula text can't change
 * while this function is only waiting on values to settle.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string[]} sheetNames
 * @param {Set<string>} excludedSheetNames
 * @param {Map<string,string>} namedRangeSheetNames
 * @returns {Map<string, number[][]>} Sheet name -> array of [row, col] pairs to watch.
 */
function buildCalculationWatchLists(spreadsheet, sheetNames, excludedSheetNames, namedRangeSheetNames) {
  const watchLists = new Map();
  sheetNames.forEach((sheetName) => {
    const sheet = spreadsheet.getSheetByName(sheetName);
    const dataRange = sheet.getDataRange();
    const numRows = dataRange.getNumRows();
    const numCols = dataRange.getNumColumns();
    const cells = [];
    if (numRows > 0 && numCols > 0) {
      const formulas = dataRange.getFormulas();
      const values = dataRange.getValues();
      for (let r = 0; r < numRows; r++) {
        getFlattenColumnIndices(formulas[r], values[r], excludedSheetNames, namedRangeSheetNames)
          .forEach((c) => cells.push([r, c]));
      }
    }
    watchLists.set(sheetName, cells);
  });
  return watchLists;
}

/**
 * Re-reads current values and returns the first watched cell still showing
 * the loading placeholder, or null if none are.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {Map<string, number[][]>} watchLists
 * @returns {{sheetName: string, a1Notation: string}|null}
 */
function findCellStillCalculating(spreadsheet, watchLists) {
  for (const [sheetName, cells] of watchLists) {
    if (cells.length === 0) continue;
    const sheet = spreadsheet.getSheetByName(sheetName);
    const values = sheet.getDataRange().getValues();
    for (const [r, c] of cells) {
      const row = values[r];
      if (row && row[c] === CALCULATION_LOADING_PLACEHOLDER) {
        return { sheetName, a1Notation: sheet.getRange(r + 1, c + 1).getA1Notation() };
      }
    }
  }
  return null;
}
