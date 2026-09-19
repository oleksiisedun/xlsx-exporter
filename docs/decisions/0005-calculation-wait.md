# 0005 — Calculation wait design

## Context

There is no public Apps Script API for spreadsheet calculation status. `SpreadsheetApp.flush()` doesn't help: it only flushes the calling script's own pending writes, and the source is never written to.

## Decision

- **Detection.** `waitForCalculationsToFinish_` (`src/CalculationWaiter.js`) polls `Range.getValues()` for the literal placeholder `"Loading..."` (`CALCULATION_LOADING_PLACEHOLDER`). The text is UI-facing and could vary by locale or Sheets version — verify with `console.log(JSON.stringify(value))` against a real slow custom function before trusting it in a non-English spreadsheet.
- **Absence of the placeholder is not proof of completion.** A cross-spreadsheet `IMPORTRANGE` read through a script can briefly report a plausible-but-stale value (e.g. `0`) while still settling, especially for wide imports or ones wrapped in `LET`/`FILTER`/`CHOOSECOLS`. A real export had two of three `IMPORTRANGE` cells silently flattened to `0` this way. The wait therefore requires **two consecutive polls** (`pollIntervalMs` apart) that both show no placeholder **and** report identical values for every watched cell; any change resets the count. Don't revert to a single-read check.
- **Scope.** Only the cells `getFlattenColumnsByRow_` would flatten are watched, not every cell: a SAFE live formula that's still calculating stays live in the export and is never read/frozen, so its transient value is irrelevant.
- **Ordering.** The wait runs *before* the Drive duplicate and orphan sweep, so a timeout never wastes a Drive copy on an export about to abort.

## Consequences

Known limitation: a formula-less cell where a user manually typed the literal `"Loading..."` is indistinguishable from a pending cell. It only surfaces as a named timeout error pointing at the exact sheet/cell, never as silent data loss.
