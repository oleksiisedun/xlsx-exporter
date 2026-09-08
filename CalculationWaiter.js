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
 * The absence of the "Loading..." placeholder is not by itself proof that a
 * cell is done: a cross-spreadsheet IMPORTRANGE read through a script (as
 * opposed to a live browser session that keeps it warm) can briefly report a
 * plausible-but-stale value — e.g. 0 — instead of the placeholder while it
 * is still settling, especially for wider imports or ones wrapped in extra
 * array steps (LET/FILTER/CHOOSECOLS...). Relying solely on the placeholder
 * check let exactly this slip through: the loop saw no "Loading..." on the
 * very first read and exited immediately, and the later, separate read in
 * `flattenUnsafeFormulas` baked in that same stale 0. To guard against this,
 * a poll only counts as settled once two consecutive reads — spaced
 * `pollIntervalMs` apart — see no "Loading..." AND report identical values
 * for every watched cell; any change between polls resets the count.
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

  let previousValues = null;
  while (true) {
    const stillCalculating = findCellStillCalculating(sourceSpreadsheet, watchLists);
    if (!stillCalculating) {
      const currentValues = captureWatchedValues(sourceSpreadsheet, watchLists);
      if (previousValues && watchedValuesEqual(previousValues, currentValues)) return;
      previousValues = currentValues;
    } else {
      previousValues = null;
    }

    if (Date.now() >= deadline) {
      const reason = stillCalculating
        ? `"${stillCalculating.sheetName}"!${stillCalculating.a1Notation} still shows "${CALCULATION_LOADING_PLACEHOLDER}"`
        : 'watched cell values were still changing between consecutive reads, never settling on two identical reads in a row';
      throw new Error(
        `waitForCalculationsToFinish: ${reason} after waiting ${timeout}ms; aborting export before an unsettled ` +
        `value bakes into the static export. Pass a larger calculationWaitTimeoutMs if this cell's source (custom ` +
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

/**
 * Reads the current value of every watched cell, keyed by sheet name, so two
 * successive polls can be compared for equality. This is what lets
 * `waitForCalculationsToFinish` catch a cell that reports a plausible but
 * still-unsettled value (e.g. a cross-spreadsheet IMPORTRANGE briefly
 * showing 0) without ever displaying the "Loading..." placeholder.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {Map<string, number[][]>} watchLists
 * @returns {Map<string, Array>} Sheet name -> cell values, in the same order as watchLists' [row, col] pairs.
 */
function captureWatchedValues(spreadsheet, watchLists) {
  const snapshot = new Map();
  for (const [sheetName, cells] of watchLists) {
    if (cells.length === 0) {
      snapshot.set(sheetName, []);
      continue;
    }
    const sheet = spreadsheet.getSheetByName(sheetName);
    const values = sheet.getDataRange().getValues();
    snapshot.set(sheetName, cells.map(([r, c]) => (values[r] ? values[r][c] : undefined)));
  }
  return snapshot;
}

/**
 * Compares two watched-value snapshots (from `captureWatchedValues`) for
 * exact equality, cell by cell. Dates are compared by timestamp since
 * `Range.getValues()` returns a distinct Date instance on every call.
 * @param {Map<string, Array>} a
 * @param {Map<string, Array>} b
 * @returns {boolean}
 */
function watchedValuesEqual(a, b) {
  for (const [sheetName, valuesA] of a) {
    const valuesB = b.get(sheetName);
    if (!valuesB || valuesA.length !== valuesB.length) return false;
    for (let i = 0; i < valuesA.length; i++) {
      const x = valuesA[i];
      const y = valuesB[i];
      const equal = x instanceof Date && y instanceof Date ? x.getTime() === y.getTime() : x === y;
      if (!equal) return false;
    }
  }
  return true;
}
