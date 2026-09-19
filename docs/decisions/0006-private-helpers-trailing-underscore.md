# 0006 — Internal helpers are private via a trailing underscore

## Context

In an Apps Script library every top-level function is exposed to consumers (`XlsxExporter.someHelper()`), and to the editor's Run menu. Only functions whose names end in `_` are private.

## Decision

The public API is `exportSpreadsheetToXlsxBlob` and `exportSpreadsheetToXlsxFile`. Every other helper ends in `_`.

## Consequences

Consumers can't couple to internals, so helpers can be renamed or removed freely. When adding a helper, name it with a trailing `_`; when adding a public entry point, document it in the README's Usage section.
