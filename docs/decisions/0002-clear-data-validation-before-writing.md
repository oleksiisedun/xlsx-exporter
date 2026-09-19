# 0002 — Clear data validation on every rewritten range before writing

## Context

Sheets enforces "reject invalid input" validation on programmatic `setValues()` calls, not just manual entry. A rule whose source list lives on a sheet we just deleted (or that simply doesn't accept the flattened value's type) makes the write throw. This was found through real testing, not anticipated up front.

## Decision

`writeFlattenedRectangles_` calls `clearDataValidations()` on each range before `setValues()`.

## Consequences

Don't remove this without re-testing against a sheet that has dropdown validation sourced from an excluded sheet. Validation on rewritten cells is lost in the export — accepted, per [0001](0001-rewrite-formula-less-cells-from-source.md).
