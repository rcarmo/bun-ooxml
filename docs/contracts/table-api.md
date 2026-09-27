# Rectangular tables

Native rectangular DOCX/PPTX tables support creation and bounded cell edits.
DOCX supports [bounded empty-row insertion and deletion](table-rows.md).
Column edits, merge/split and PPTX row mutations are not implemented. Existing merged tables can be inspected only when grid mapping
is unambiguous; unsafe edits refuse before mutation. No formatting reconstruction
outside the selected table/cell.

DOCX: `Document.addTable(rows:number, columns:number): Table`, `Document.tables: Table[]`.
Table `rows:number`, `columns:number`, `cell(row,col): TableCell`,
`tryCell(row,col): TableCell | null` and
[`rowTexts(row): string[]`](row-texts.md) for detached cell text in column order. TableCell `text:string`
(get/set), supports direct simple paragraphs/runs only; setter may replace the cell
text payload while retaining the first paragraph and run properties. Empty cells valid. Added table goes before sectPr,
valid tblPr/tblGrid/tr/tc/tcPr/p structure, bounded dimensions, table handles refuse
on structural document edits and cell handles refuse after any mutation. Newlines
normalise CRLF/CR and become paragraphs. The first paragraph/run properties are
retained with their namespace declarations. `TableCell.directProperties()` and
`setProperties()` provide [bounded direct cell formatting](cell-properties.md)
without replacing text or changing the grid. `Table.isRowHeader(row)` and
`setRowHeader(row, boolean | null)` inspect and edit [direct row-header markers](row-header.md).
They do not predict rendered page repetition. `Table.styleId` and `setStyle(id | null)`
read or edit a [direct table-style reference](table-style.md) without creating or
resolving the style definition. Full style inheritance and merging are unsupported.

### Nullable DOCX cell access

`tryCell(row, column)` returns an ordinary checked `TableCell` for a supported
in-range coordinate, or literal `null` for an out-of-range integer coordinate.
Both arguments must be safe integers; strings, fractions, nonfinite and unsafe
numbers throw `RangeError` without coercion. `cell()` retains its throwing bounds
policy. The nullable method does not catch and suppress errors from `cell()`.

Stale tables and unsupported grid mappings refuse before returning `null`,
including after external main-document changes. A merged target still refuses;
an out-of-range coordinate on an otherwise mappable merged table returns `null`.
A nested or otherwise unsupported cell can yield a handle, but its text access
still refuses. Returned cell handles use the existing stale-target and mutation
checks. No nullable access is added to PPTX.

The shared `@id-docx-go-table-cell-access` case checks nine nonnil cells and four
literal null boundary results. Its binding uses `tryCell` directly. Native tests
separately cover coordinate validation, stale/merged/unsupported refusals, exact
unchanged bytes, checked cell editing and path save/reopen with unrelated-member
custody. Seven partial mappings record those limits; only the shared wrapper
selects a case key.

PPTX: `Slide.addTable(rows,columns,{x,y,width,height}): Table` with integer EMU geometry.
`Slide.tables: Table[]`, same rows/columns/cell text API. Native graphicFrame/a:tbl,
unique cNvPr id, grid column widths sum width and row heights sum height, xfrm geometry,
valid table properties and trailing p/r instructions. Native reader table/text APIs
must not corrupt existing shape identities. Bounds max10000 cells, positive rows/cols,
positive widths/heights, finite safe integer geometry. Zero x/y allowed. PPTX table
and cell handles become stale after a slide mutation; reacquire from Slide.tables.

Both: zero-based indices, typed range errors, unchanged saved bytes on refusal;
XML escaping and boundary whitespace; saved/reopened table content and opaque
parts verified. Reacquire stale handles after edits; they are checked against
fingerprints or revision counters. The shared specification defines table behaviour in
`references/fixtures-ooxml/workflows/docx/tables.feature` and
`references/fixtures-ooxml/workflows/pptx/tables.feature`.
