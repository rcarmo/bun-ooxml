# Direct paragraph run formatting

`Paragraph.setRunFormatting(patch)` applies direct Boolean,
[font-size](font-size.md), [Latin font-name](run-font-name.md) and
[scalar appearance](run-appearance.md) overrides to every supported direct run in one paragraph. It returns `{ changedRuns: number }`.
`RunFormattingPatch` and `RunFormattingReceipt` are exported from the root and
`./docx` entrypoints.

```ts
const paragraph = document.paragraphs[0]!;
paragraph.setRunFormatting({ bold: true, italic: false });
// Reacquire after a real mutation. Null removes a direct override.
document.paragraphs[0]!.setRunFormatting({ bold: null });
```

Boolean fields are `bold`, `italic`, `strike`, `doubleStrike`, `caps`, `smallCaps`,
`outline`, `shadow`, `emboss`, `imprint` and `vanish`. They accept true, false, null
or undefined. True writes on; false writes explicit off. `fontSizePt` accepts positive numbers exactly representable
in half-points. For all fields, null removes the direct property and undefined
leaves it unchanged. `Paragraph.directRunFlags()` returns detached eleven-field
Boolean/null objects in run order; null means no direct property. It does not
apply inherited toggle properties or defaults. `DirectRunFlags` is exported from
the root and `./docx` entrypoints.

This API does not compute inherited formatting or alter paragraph/character
styles. It applies to the whole paragraph, including paragraphs in table cells;
substring formatting and run splitting are unsupported.

## Preservation and handles

Text and all package parts outside the main document remain byte-identical.
Unrelated run-property elements retain their original bytes. Selected formatting
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

Patches must contain own plain data properties. Getters, symbols, unknown keys and
custom prototypes refuse without invoking accessor properties.

The editor refuses these same-run direct-element pairs: `strike`/`doubleStrike`,
`caps`/`smallCaps`, `emboss`/`imprint`, `emboss`/`outline`, `imprint`/`outline`,
`emboss`/`shadow` and `imprint`/`shadow`. This conservative check uses presence even
when a value is false. Remove the old property with null when switching to its
alternative. Already-conflicting input refuses; this API does not repair it.

The API accepts only plain direct text runs and a conservative set of ordered
leaf run properties. Fields, tracked content, controls, mixed lexical content,
property revisions, unknown/duplicated/out-of-order properties and malformed or
misqualified Boolean values refuse. Font-size edits and reads also reject
malformed, misqualified or unsupported direct size values. Empty paragraphs report zero changed runs;
nonempty formatting on self-closing empty runs refuses.

All linked settings parts must be inspectable and unprotected. Missing or unknown
protection enforcement refuses; only explicit off values are accepted. External
settings refuse. A stale paragraph or direct underlying document change refuses
rather than reusing old offsets. Error codes are `docx-format-argument`,
`docx-format-unsupported`, `docx-format-protected`, `docx-format-unsafe`, or the
existing stale/unsupported paragraph codes.

## Specification and coverage

ECMA-376 Part 1, fifth edition (October 2016), defines `caps` in §17.3.2.5,
`dstrike` in §17.3.2.9, `emboss` in §17.3.2.13, `imprint` in §17.3.2.18,
`outline` in §17.3.2.23, `shadow` in §17.3.2.31, `smallCaps` in §17.3.2.33,
`strike` in §17.3.2.37 and `vanish` in §17.3.2.41. The conflict checks follow the
same-run exclusions in those clauses. The complete PDF is in the
[shared specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).

The shared five-row bold/italic/strike outline in `docx/document-model.feature`
executes in Bun, alongside the existing `docx/run-formatting.feature` cases. The
separate Go all-eight-effects getter case stays planned: it sets mutually exclusive
effects on one run and checks in-memory values without saving XML. Individual
valid effects have native save/reopen tests and one Open XML SDK Office2019
schema-validated sample. Word rendering, complex-script meanings and broad
style-inheritance compatibility have not been tested.
