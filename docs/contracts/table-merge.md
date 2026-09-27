# Horizontal Word cell merges

`Table.mergeRowCells(row, firstColumn, lastColumn)` merges an inclusive horizontal
range and returns a fresh table handle. Coordinates are zero-based safe integers;
the range must contain at least two cells. [Single-column vertical merges](table-vertical-merge.md)
have a separate method. Splitting is unsupported.

```ts
const table = document.tables[0]!;
// Cells (0, 1) and (0, 2) must be structurally empty and unformatted.
const merged = table.mergeRowCells(0, 0, 2);
```

The first cell retains its paragraphs, text and formatting. The writer sums the
selected grid widths, updates its `tcW`, writes `gridSpan`, and removes the other
selected cell elements. It leaves the table grid and unrelated rows/members
unchanged. It never concatenates text or discards populated or formatted cells.

## Admission

The whole table must have one explicit grid, at most 100 rows and 100 columns,
and one unmerged cell per grid column. Cell widths must be explicit `dxa` values
matching grid widths. Individual widths and the merged total cannot exceed
31,680 twips. Existing spans/merges, nested tables, property revisions, ambiguous
metadata, lexical barriers, section-bearing cell paragraphs and unsupported run
content refuse before mutation.

Each absorbed cell must contain only `tcPr` with a plain `tcW`, followed by one
structurally empty paragraph. Text, whitespace text, empty runs, paragraph or cell
formatting, and non-namespace attributes on discarded containers refuse. The
first cell may retain supported direct formatting and plain text paragraphs.
Protected documents and external settings refuse. Singleton, reversed, invalid
or out-of-range coordinates throw `RangeError`.

## Persistence and handles

The writer validates the exact retained paragraph sequence and table dimensions,
then writes and serializes inside the package transaction. A failure restores
bytes and keeps prior handles usable. Success invalidates all prior table,
paragraph and cell handles; use the returned table or reacquire from `Document`.
UTF-8 BOM and UTF-16 encoding survive the write.

The returned table still refuses `cell()`/`tryCell()` at merged coordinates.
Existing row editing and another merge on that table also refuse. Unmerged cells
can still be read. Editing merged content and repeated merges require further
implementation.

## Evidence and shared limits

Twelve native tests cover saved/reopened output, exact grid/sibling/opaque bytes,
format retention, refusals, fault rollback, encoding and stale handles. The
optional `make office-oracles` gate validates one authored horizontal merge with
the Open XML SDK and reopens it natively. It does not test rendered geometry.

The shared `@id-docx-go-table-merge-properties` case remains planned: it requires
individual span and vertical-merge setters/getters, with a different topology
policy. Shared v0.29 adds seven [physical merge contracts](../../references/fixtures-ooxml/contracts/table-merging.md)
with 26 cases, executed by `tests/acceptance/table-merging.ts`. They check saved
geometry, exact retained paragraph order and sibling/source bytes, content-loss
refusals, reached fault rollback, three encoding markers and stale handles.
UTF-16BE now has explicit outcome coverage alongside UTF-16LE and UTF-8 BOM.

Fifteen partial mappings cover twelve native declarations and three outcome
wrappers. Only the positive shared wrapper selects the 26 case keys. The
historical setter-profile mapping is an inventory-only guard with no case keys;
Python and Go gain no execution or parity credit from these Bun results.
