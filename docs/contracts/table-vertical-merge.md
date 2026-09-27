# Vertical Word cell merges

`Table.mergeColumnCells(column, firstRow, lastRow)` merges an inclusive range in
one grid column and returns a fresh table handle. All three coordinates must be
safe integers; at least two consecutive rows are required.

```ts
const table = document.tables[0]!;
// Cells (1, 1) and (2, 1) must have only a width and one empty paragraph.
const merged = table.mergeColumnCells(1, 0, 2);
```

The first cell keeps its content and formatting. The writer inserts `vMerge`
with explicit `restart` in the owner and `continue` in every following selected
cell. It preserves every cell, paragraph, width, grid column and unrelated member.
Removing only the inserted markers reproduces the decoded source XML exactly.
ECMA-376 Part 1 §17.4.84 defines these markers and requires matching grid columns.

## Admission and handles

The operation uses the [horizontal merge admission checks](table-merge.md#admission):
one explicit grid, at most 100 rows/columns, matching explicit cell/grid widths,
plain cell paragraphs and no existing merges anywhere in the table. It also
requires every continuation cell to contain only a width and one structurally
empty paragraph. Text (including spaces), empty runs, paragraph or cell formatting
and lexical barriers refuse; the writer does not hide content under a merge.

Invalid, reversed, singleton or out-of-range coordinates throw `RangeError`
without coercion. Protected documents, external settings, revised or ambiguous
properties, nested cells and stale handles refuse before writing. Writes and
serialization are transactional; failed publication restores bytes and preserves
existing handles. Success invalidates old table, paragraph and cell handles.

The returned table still refuses editable `cell()`/`tryCell()` access at merged
coordinates. Row edits and subsequent horizontal or vertical merges also refuse.
Other cells remain readable. Splitting, multi-column vertical ranges, repeated
merges, editing merged content and inherited layout are unsupported.

## Verification and limits

Twelve native tests check path save/reopen, exact sibling XML, owner formatting,
all continuation markers, retained paragraph order, unselected tables and opaque
parts, atomic refusals, reached write/serialization faults and stale handles.
UTF-8 BOM and aliased UTF-16LE/BE inputs retain their markers and namespace meaning.
The optional Office oracle validates one authored sample with the Open XML SDK
and reopens it natively; rendered geometry is untested.

The native declarations are unmapped pending central physical vertical-merge
contracts. The existing shared horizontal contract and individual span/vertical
setter profile remain unchanged. No shared execution or other-consumer parity
credit is granted by this addition.
