# Insert and delete Word table rows

`Table.insertRow(index)`, `appendRow()` and `deleteRow(index)` return a fresh table
handle. Row indexes are zero-based. Insertion accepts positions from zero through
the current row count; deletion accepts only existing rows. At least one row must
remain after deletion.

```ts
let table = document.addTable(2, 3);
table = table.appendRow();       // three rows
table = table.insertRow(1);      // four rows
table.cell(1, 0).text = 'New';
table = table.deleteRow(2);      // three rows
```

Each successful operation expires every previous paragraph, span, cell and table
handle in the document, including handles to unrelated tables. Use the returned
table or reacquire handles from the document. A failed operation retains handles
and package bytes.

## New rows and preservation

New rows have one empty paragraph per cell. Their preferred cell widths come from
the existing explicit grid, expressed as `dxa`. They do not copy neighbouring
header flags, row heights, cell properties or run formatting. Table styles can
still affect their appearance. Existing row and grid fragments retain their
source bytes; namespace bindings on the authored row prevent alias capture.

Deletion removes the selected row and all its cells and paragraphs. The table
grid, remaining rows, other body content and unrelated package parts remain
unchanged. Inserting then deleting the new row restores the original main XML in
the tested explicit-grid example. Row positions change, so conditional table-style
formatting or contiguous header behaviour can change without a markup edit to a
surviving row. Effective styles and rendered row heights are not computed.

## Supported tables and refusals

The table must be top-level, rectangular, with 1–100 rows and columns and a single
explicit grid. Every grid column must have an integer width from 0 through 31680
twips. These are editor limits; a missing or inferred width is not generated.
Deletion applies the same grid checks as insertion.

All rows are checked before either operation. Merged or nested cells, grid
omissions, fields, revisions, section-bearing cell paragraphs, lexical barriers
and unsupported extension attributes refuse. Row and cell metadata must pass the
existing bounded property checks. Paragraphs must contain plain direct runs;
text validation also applies the paragraph setter's 1 Mi UTF-16-unit and
no-tab/line-break limits. This validation does not rewrite text. Unsupported
content in any row prevents editing the table, even when that row would remain.

The target must match the current table snapshot and exact main-part XML.
Protection checks run before mutation. The complete prospective document must
parse, retain the expected paragraph sequence and have exactly the intended
one-row dimension change. Package write and serialization run inside a
transaction before model versions change. These checks do not validate complete
table semantics. The XML editor also limits complete output to 8 Mi UTF-16 units;
this is not a peak-memory guarantee.

## Specification and checks

ECMA-376 Part 1, fifth edition, §17.4.78 defines `tr`; §17.4.48 defines the table
grid. The [specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md)
contains the source PDFs. The shared row-count case checks the 3/4/3 count sequence
and invalid deletion on a 2x3 table. Native tests separately check saved text/order,
widths, exact fragments, namespace aliases, UTF-16 and rollback. An authored sample
passes Open XML SDK validation, and the Office smoke test includes all three row
operations. General table layout, merging, column edits and PPTX row mutation are
outside this API.
