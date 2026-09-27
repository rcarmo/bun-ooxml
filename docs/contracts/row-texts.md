# Read Word table-row text

`Table.rowTexts(index)` returns a new array containing each cell's text in column
order. Indexes are zero-based; `table.rowTexts(0)` reads the first row. Changing
the returned array does not change the document.

```ts
const firstRow = document.tables[0]!.rowTexts(0);
```

Values use the same plain-text policy as `TableCell.text`: direct run text is
concatenated, empty cells return an empty string, and multiple cell paragraphs
are joined with newlines. Entities are decoded. Lexical comments inside a text
leaf contribute no text and remain untouched in the package. This is not a
rendered-text, field-evaluation or whitespace-layout API.

The reader resolves the current table snapshot, checks the exact main-part XML,
and checks every selected cell before returning. Invalid row indexes throw
`RangeError`; they do not return an empty array or null. Merged cells, unresolved
coordinates, unsupported table topology and selected cells containing nested
tables or fields refuse. An unsupported cell in another row does not by itself
prevent reading a supported row. Table-wide grid errors still refuse.

The existing model can infer columns when a table lacks an explicit grid; this
reader follows that model and does not add schema validation or reinterpret its
numeric grid parsing. Row mutation uses a stricter explicit-grid contract.

Inspection does not write package parts, expire handles or block on document
protection. A held table follows supported cell edits, but structural edits expire
it. Out-of-band main-part changes refuse instead of returning cached values.

The shared first-row scenario checks four in-memory cell getters and exactly two
first-row values. Native tests separately cover detached results, stale sources,
protected reads, path save/reopen and aliased UTF-16 packages. The out-of-range
cell-access scenario remains unsupported because its null-return policy differs
from Bun's throwing API. No new Word markup is authored by this reader.
