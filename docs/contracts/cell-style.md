# Existing XLSX cell styles

`Worksheet.setCellStyle(address, styleIndex | null)` returns `{changed: 0 | 1}`.
The cell must already exist. A numeric index selects an existing `cellXfs` entry;
zero writes an explicit `s="0"`. Null removes only the direct `s` attribute and
does not require a styles part. Row/column defaults may then affect the cell.
This API does not compute effective formatting or create cells/style definitions.

Selection requires one exact internal styles relationship, the expected content
type and SpreadsheetML root, unambiguous relevant collections and correct declared
counts. The selected cellXf and its optional base must resolve font/fill/border
references. Custom number-format IDs require definitions. Unrelated cellXfs and
full formatting schemas are not validated. Existing index spelling survives a
numeric no-op. A malformed current `s` refuses, including removal.

Only the target cell's style attribute changes. Values, formulas, cached answers,
other attributes, worksheet content and all other package payloads stay unchanged.
Formula/cache topology is not rewritten, including shared/array formula metadata.
Serialisation occurs inside the rollback boundary; cached cell state updates only
after success. UTF-16 encoding and BOM survive style-only writes.

The writer refuses any sheet/workbook protection element conservatively, including
disabled protection, and ambiguous worksheet aliases to the same XML part. It
checks raw bytes for the cached main document, workbook/root relationships,
worksheets and shared strings before mutation. Out-of-band changes require reopening
the workbook. Style definitions are read live. Existing worksheet facades keep
working after a supported mutation; returned Cell objects remain detached values.

Tests cover 27 shared cases plus encoding, lexical preservation, dependency
validation, model aliases, serialisation rollback and disk save/reopen. Native
readback does not establish Excel rendering, computed number formats, conditional
formatting, calculation or full style parity.

Canonical feature: `workflows/xlsx/cell-style.feature` in the shared reference.
Bindings/tests: `tests/acceptance/cell-style.ts`, `tests/unit/xlsx-cell-style.test.ts`.
