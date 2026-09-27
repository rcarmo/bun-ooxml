# Direct Word table-row header markers

`Table.isRowHeader(row)` reads the direct `w:tblHeader` marker on a zero-based
row. It returns false when the marker is absent or explicitly disabled.
`Table.setRowHeader(row, value)` accepts a Boolean or `null` and returns
`{ changed: 0 | 1 }`.

```ts
const table = document.addTable(3, 2);
table.cell(0, 0).text = 'Name';
table.cell(0, 1).text = 'Value';
table.setRowHeader(0, true);
console.log(table.isRowHeader(0)); // true
```

`true` writes an enabled marker; `false` writes an explicit disabled marker.
`null` removes the marker. Thus setting false on an unmarked row changes XML,
while setting null on that row does not. Existing Boolean spellings are retained
on same-value calls. New markers use `on` and `off`.

A changed write expires paragraph, span and cell handles. The table handle remains
usable; reacquire its cells. Validated no-ops preserve archive bytes and handles.
Writes, including no-ops, check document protection. Reads and writes reject stale
tables or document XML changed outside the model.

## Supported targets

The table must have an explicit rectangular grid. The selected row must contain
plain, unmerged cells and at most one leading `trPr`. Duplicate properties,
revisions, unknown or misqualified row properties, decorated header markers and
lexical barriers refuse before mutation. `trPrBase` properties are an unordered
choice; their relative order is retained. Other permitted row metadata is checked
structurally and preserved, not fully interpreted or schema-validated.

Edits change only the marker or its property container. Existing cell fragments,
sibling rows and unrelated package parts retain their bytes. Namespace aliases and
UTF-16 document encoding survive save/reopen. Prospective output must parse and
retain paragraph text and table dimensions. Package write and serialization are
transactional; failure retains both bytes and handles. The XML editor's complete
output limit also applies.

## Specification and compatibility

ECMA-376 Part 1, fifth edition, §17.4.49 defines `tblHeader`; §17.17.4 defines its
Boolean values. Part 4's Transitional schema also declares it as `CT_OnOff`.
The reader accepts absent `val`, `true`, `false`, `1`, `0`, `on` and `off`.
Open XML SDK 3.5.1 applies a narrower `OnOffOnlyType` here and rejects numeric
values, so authoring uses the compatible `on`/`off` spellings. Same-value calls do
not rewrite existing numeric markers. See the [complete specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).

This API reads direct markup, not table-style inheritance or rendered pagination.
Word repeats only a contiguous block of marked rows starting with the first row;
it can ignore an isolated marker on a later row. We permit that direct marker and
do not report it as effective repetition. SDK validation checks authored on/off
markers; it does not establish multi-page header rendering.
