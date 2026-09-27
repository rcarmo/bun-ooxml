# Word template inventory

`Document.inspectTemplate()` reads concrete content from the main document body
and its plain rectangular top-level tables. It returns a detached, deeply frozen
`TemplateInspection` and leaves archive bytes and existing handles unchanged.

```ts
const document = await Document.open('template.docx');
const report = document.inspectTemplate();
for (const placeholder of report.placeholders) {
  console.log(placeholder.paragraphIndex, placeholder.name, placeholder.start);
}
```

The report contains:

- `scope: "main-body-and-tables"` and exact counts;
- `paragraphs`: document-order text, direct style ID or `null`, direct outline
  level or `null`, body index, and table/row/column/paragraph coordinates for cells;
- `sections`: direct body paragraphs with outline levels 0–8, reported as levels
  1–9 with their paragraph index, body index and text;
- `tables`: dimensions and row-major cell text (cell paragraphs joined with `\n`);
- `placeholders`: literal angle-delimited occurrences with paragraph/body indexes,
  original text, name and start/end offsets in decoded UTF-16 units.

Body indexes count direct paragraphs and tables, excluding final section
properties. Paragraph indexes include plain cell paragraphs. Placeholder names
are nonnested, nonblank, at most 256 UTF-16 units and contain no tab/line break.
Occurrences can cross plain runs but never paragraph boundaries. Repeated names
remain separate occurrences. Square-bracket tokens such as `[TBD]` are not scanned.
The scanner is shared with `inspectBodyMap`; that API keeps its body-only scope.

## Limits and refusal

The report has no SOW, guidance-colour or staffing inference, cached metadata,
effective formatting, inherited headings or rendering prediction. Style IDs are
literal direct values; section classification uses only direct body outline
levels. Headers, footers, notes, comments and other stories are outside the named
scope. No external content is fetched.

The reader accepts Transitional Word body XML with plain paragraphs and direct
runs, plus rectangular, unmerged tables with explicit grids and matching explicit
twip cell widths. Row grid offsets must be absent or lexically valid zero values.
Fields, hyperlinks,
controls, revisions, nested/merged tables, malformed selected properties, missing
or mismatched grids, lexical barriers and unknown body structure refuse the whole
inspection. A stale wrapper refuses instead of inspecting an old snapshot.
Protection permits reading; it grants no editing permission.

Bounds are 10,000 total paragraphs, 10,000 placeholder occurrences, 1,000 tables,
and 100 rows/columns per table, in addition to the XML parser limits. Exceeding a
bound throws `docx-template-limit`; unsupported table input throws
`docx-template-unsupported` or an existing paragraph/cell admission error. There
is no partial-success or truncated result. Outputs remain valid snapshots after
later document edits and expose no editable anchor handles.

Fourteen native tests cover concrete output, path save/reopen, unchanged bytes,
immutable snapshots, UTF-16BE aliases, bounds and late refusal. These tests are
unmapped pending central concrete-output contracts. Existing template response
shape, status, cache and metadata profiles remain planned; returning an object
does not satisfy their distinct tool/cache APIs or establish semantic analysis.
