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
 * Blocks until every cell that `flattenUnsafeFormulas_` (SpreadsheetDuplicator.js)
 * would bake into the export as a static value — i.e. every cell
 * `getFlattenColumnsByRow_` (FormulaClassifier.js) selects — has finished
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
 * `flattenUnsafeFormulas_` baked in that same stale 0. To guard against this,
 * a poll only counts as settled once two consecutive reads — spaced
 * `pollIntervalMs` apart — see no "Loading..." AND report identical values
 * for every watched cell; any change between polls resets the count.
 *
 * The watched cell set is re-derived from fresh formulas and values on every
 * poll (see `readFlattenCellsSnapshot_`), so cells a still-loading spilling
 * formula fills in later are picked up too, and a spill changing size counts
 * as a change.
 *
 * Pass `timeoutMs: 0` to opt out entirely — returns immediately, no reads.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} sourceSpreadsheet
 * @param {string[]} includedSheetNames
 * @param {Set<string>} excludedSheetNames
 * @param {Map<string,string>} namedRangeSheetNames
 * @param {DeletedColumns} deletedColumns
 * @param {number} [timeoutMs] - Max total wait before throwing. Defaults to 120000 (2 min). Uses `??`, not `||`, so 0 is a meaningful opt-out.
 * @param {number} [pollIntervalMs] - Delay between re-checks. Defaults to 3000 (3s); floored at 250ms.
 * @returns {void}
 */
function waitForCalculationsToFinish_(sourceSpreadsheet, includedSheetNames, excludedSheetNames, namedRangeSheetNames, deletedColumns, timeoutMs, pollIntervalMs) {
  const timeout = timeoutMs ?? 120000;
  if (timeout === 0) return;
  const interval = Math.max(pollIntervalMs ?? 3000, 250);

  const deadline = Date.now() + timeout;

  /** @type {FlattenCellsSnapshot|null} */
  let previous = null;
  while (true) {
    const current = readFlattenCellsSnapshot_(sourceSpreadsheet, includedSheetNames, excludedSheetNames, namedRangeSheetNames, deletedColumns);
    const stillCalculating = current.loadingCell;
    if (!stillCalculating) {
      if (previous && watchedValuesEqual_(previous.positions, current.positions) && watchedValuesEqual_(previous.values, current.values)) return;
      previous = current;
    } else {
      previous = null;
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
 * @typedef {Object} FlattenCellsSnapshot
 * @property {Map<string, string[]>} positions - Sheet name -> "row:col" of every cell `getFlattenColumnsByRow_` selects, in scan order.
 * @property {Map<string, any[]>} values - Sheet name -> those cells' values, parallel to `positions`.
 * @property {{sheetName: string, a1Notation: string}|null} loadingCell - First selected cell still showing the loading placeholder, or null.
 */

/**
 * Re-reads formulas and values from scratch and returns exactly the cells
 * `getFlattenColumnsByRow_` would freeze into the export right now, with
 * their values. The selection is recomputed on every poll, never cached: a
 * spilling formula (IMPORTRANGE/FILTER/...) that is still loading holds a
 * single cell, and the cells it later spills into are formula-less values
 * that only join the selection once they appear. A selection frozen from the
 * first read would never watch them, so they could still be changing — or be
 * a stale 0 — when the export reads them. Comparing positions as well as
 * values makes a spill that grows or shrinks between polls count as a change.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} spreadsheet
 * @param {string[]} sheetNames
 * @param {Set<string>} excludedSheetNames
 * @param {Map<string,string>} namedRangeSheetNames
 * @param {DeletedColumns} deletedColumns
 * @returns {FlattenCellsSnapshot}
 */
function readFlattenCellsSnapshot_(spreadsheet, sheetNames, excludedSheetNames, namedRangeSheetNames, deletedColumns) {
  /** @type {FlattenCellsSnapshot} */
  const snapshot = { positions: new Map(), values: new Map(), loadingCell: null };
  sheetNames.forEach((sheetName) => {
    const sheet = getRequiredSheet_(spreadsheet, sheetName);
    const dataRange = sheet.getDataRange();
    /** @type {string[]} */
    const positions = [];
    /** @type {any[]} */
    const cellValues = [];
    if (dataRange.getNumRows() > 0 && dataRange.getNumColumns() > 0) {
      const values = dataRange.getValues();
      getFlattenColumnsByRow_(dataRange.getFormulas(), values, sheetName, excludedSheetNames, namedRangeSheetNames, deletedColumns)
        .forEach((cols, r) => cols.forEach((c) => {
          positions.push(`${r}:${c}`);
          cellValues.push(values[r][c]);
          if (!snapshot.loadingCell && values[r][c] === CALCULATION_LOADING_PLACEHOLDER) {
            snapshot.loadingCell = { sheetName, a1Notation: sheet.getRange(r + 1, c + 1).getA1Notation() };
          }
        }));
    }
    snapshot.positions.set(sheetName, positions);
    snapshot.values.set(sheetName, cellValues);
  });
  return snapshot;
}

/**
 * Compares two watched-value snapshots (from `captureWatchedValues_`) for
 * exact equality, cell by cell. Dates are compared by timestamp since
 * `Range.getValues()` returns a distinct Date instance on every call.
 * @param {Map<string, any[]>} a
 * @param {Map<string, any[]>} b
 * @returns {boolean}
 */
function watchedValuesEqual_(a, b) {
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
