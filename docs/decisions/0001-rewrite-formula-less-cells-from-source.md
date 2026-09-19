# 0001 — Rewrite formula-less cells from the source (only on sheets with an unsafe formula)

## Context

`ARRAYFORMULA`/`QUERY`/`SORT`/`IMPORTRANGE`-style formulas store their formula text only in the top-left anchor cell. Every other cell they visually fill has an empty formula string (`Range.getFormulas()` returns `""`) and no real stored content. The instant the anchor is flattened to a literal, Sheets drops the spill and those cells go blank. Apps Script's API cannot report a formula's spill boundaries, and a spill artifact is indistinguishable from a manually typed literal.

## Decision

- Only cells holding their own **SAFE** live formula are left untouched.
- On a sheet that contains at least one **UNSAFE** formula, every other non-blank cell is rewritten from the source's captured value. A real literal is rewritten with its own unchanged value (a harmless no-op); a spill artifact is correctly preserved.
- A sheet with **no** unsafe formula is skipped entirely: a spill can never cross sheets, so with no anchor to flatten there is nothing to protect. This avoids needlessly re-parsing every literal in such a sheet through `setValues()`.
- `getFlattenColumnsByRow_` in `src/FormulaClassifier.js` is the single source of truth for this selection, shared by `flattenUnsafeFormulas_` (which freezes the cells) and `buildCalculationWatchLists_` (which watches them for pending calculation, see [0005](0005-calculation-wait.md)). Don't reimplement the scoping inline in either place.

## Consequences

- Correctness (no data loss) is prioritized over minimizing write volume or preserving data-validation fidelity on cells incidentally caught by the rewrite.
- Known residual risk (verify on a scratch sheet, not yet tested): `setValues()` treats strings like user input, so a text literal such as `'=1+1` or `'00123` in a sheet that *does* get rewritten may be re-interpreted as a formula/number. This also applies to string results of flattened formulas.
- Writes are batched: per-row column runs are merged into rectangles (`mergeColumnsIntoRectangles_`) so API calls scale with distinct blocks, not rows × runs, keeping large sheets inside the 6-minute execution limit.
