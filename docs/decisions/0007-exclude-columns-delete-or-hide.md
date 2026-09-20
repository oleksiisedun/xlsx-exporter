# 0007 — `excludeColumns`: delete or hide, flatten anything that reads a deleted column

## Context

Sheet-level exclusion already exists. Removing individual columns from a kept sheet has the same hazard (a live formula reading removed data breaks in the export) plus two of its own: Sheets doesn't always raise `#REF!` (a range that merely *overlaps* a deleted column silently shrinks and returns a different value), and the flatten step addresses cells by their original column position, so positions must not shift until flattening is done.

## Decision

- `excludeColumns: { 'Sheet': ['C', 'F:H'] }` plus `excludeColumnsMode: 'delete' | 'hide'` (default `'delete'`). Columns are removed from the **duplicate** only, and **after** `flattenUnsafeFormulas_`, right to left.
- **Delete mode:** any formula (on any included sheet) with a reference that overlaps a deleted column is UNSAFE and flattened from the source: qualified or unqualified, `A1`, `A1:B2`, `C:C`, `$C$2`, whole-row `2:2` (spans every column), and named ranges overlapping a deleted column. Overlap, not just full containment, because of the silent range shrink above.
- **Hide mode:** nothing is flattened extra. Hidden columns keep their data and references, so nothing breaks. The data is still in the file, which is the whole trade-off against delete.
- Cells inside deleted columns are never flattened or watched by the calculation wait. A formula in a deleted column still counts as "an unsafe formula on this sheet" for [0001](0001-rewrite-formula-less-cells-from-source.md): deleting its column drops any spill it produced into surviving columns, so those cells need rewriting.
- The reference regex (`CELL_REFERENCE_RE` in `FormulaParser.js`) errs toward over-matching; a false positive only turns a live formula into its value.

## Consequences

- Charts, pivot tables, data validation and conditional-format rules pointing at a deleted column are not detected (same limitation as excluded sheets). Use `'hide'` if something like that depends on the column.
- A sheet can't have all its columns deleted (Sheets refuses); this is checked up front. Column specs are validated before any Drive work.
- Untested against real Sheets so far (no harness): verify deletion with frozen columns, merged cells and filters on a scratch sheet.
