# Direct paragraph run formatting

`Paragraph.setRunFormatting(patch)` applies direct bold/italic overrides to every
supported direct run in one paragraph. It returns `{ changedRuns: number }`.
`RunFormattingPatch` and `RunFormattingReceipt` are exported from the root and
`./docx` entrypoints.

```ts
const paragraph = document.paragraphs[0]!;
paragraph.setRunFormatting({ bold: true, italic: false });
// Reacquire after a real mutation. Null removes a direct override.
document.paragraphs[0]!.setRunFormatting({ bold: null });
```

Each field accepts true, false, null or undefined. True writes on; false writes
explicit off; null removes the direct property; undefined leaves it unchanged.
This API does not compute inherited formatting or alter paragraph/character
styles. It applies to the whole paragraph, including paragraphs in table cells;
substring formatting and run splitting are unsupported.

## Preservation and handles

Text and all package parts outside the main document remain byte-identical.
Unrelated run-property elements retain their original bytes. Selected bold/italic
elements may be replaced by namespace-qualified elements; additions follow Word
run-property order. Source encoding/BOM and namespace identity survive save/reopen.

A same-state or empty patch returns zero changed runs and retains exact archive
bytes and handles. A real change invalidates paragraph, span and cell snapshots.
Table handles remain usable because the grid is unchanged; reacquire their cells.
Fresh spans can replace text after formatting without losing the direct properties.

All runs preflight before mutation. The package edit and serialization validation
run inside a transaction; the model version and snapshots update only after
success. A thrown serialization error preserves the package and existing handles.

## Refusals

The API accepts only plain direct text runs and a conservative set of ordered
leaf run properties. Fields, tracked content, controls, mixed lexical content,
property revisions, unknown/duplicated/out-of-order properties and malformed or
misqualified Boolean values refuse. Empty paragraphs report zero changed runs;
nonempty formatting on self-closing empty runs refuses.

All linked settings parts must be inspectable and unprotected. Missing or unknown
protection enforcement refuses; only explicit off values are accepted. External
settings refuse. A stale paragraph or direct underlying document change refuses
rather than reusing old offsets. Error codes are `docx-format-argument`,
`docx-format-unsupported`, `docx-format-protected`, `docx-format-unsafe`, or the
existing stale/unsupported paragraph codes.

The shared specification defines expected behaviour in
`workflows/docx/run-formatting.feature`. Independent Word rendering and broad
schema/style inheritance compatibility have not been tested.
