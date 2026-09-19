# 0003 — `IMPORTRANGE` values always come from the source spreadsheet

## Context

A Drive copy is a new file ID and starts without `IMPORTRANGE` authorization, so the duplicate can't evaluate the formula.

## Decision

`IMPORTRANGE` is always classified UNSAFE, so its cells are always flattened using the value read from the already-authorized source. The duplicate's own (unauthorized) evaluation of that formula is never read.

## Consequences

The same "read from source, write to duplicate — never the reverse" rule keeps formulas that reference excluded sheets correct instead of `#REF!`. The source is never mutated.
