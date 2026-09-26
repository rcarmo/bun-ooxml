# Paragraph-style authoring

`Document.addParagraphStyle(styleId, options)` adds a custom paragraph style and
returns `{styleId, partName}`. `options.name` is required; `basedOn`, `bold` and
`italic` are optional. Boolean flags are explicit on/off values. An omitted flag
adds no direct override. The API does not calculate inherited formatting.

The operation creates a styles part, relationship and content-type override when
no registry is linked. Orphan members are never adopted or overwritten. With an
existing registry, only that XML part changes. Existing definitions, document XML,
relationships, content types and other payloads remain byte-identical. UTF-16
encoding and BOM are retained when appending to an existing registry.

Creation requires one internal styles relationship with the expected content type
and Word namespace root. All registry style IDs must be qualified and unique;
this is stricter than the selected-ID lookup used by `Paragraph.setStyle()`.
The new ID must be unused. A requested base and its entire `basedOn` chain must
resolve uniquely to paragraph styles without cycles or malformed links. Existing
styles outside that chain are not fully schema-validated.

The writer refuses active/ambiguous document protection, external settings,
stylesWithEffects dual registries, unsupported registry children/order or lexical
barriers, stale document XML, invalid strings/options and option accessors.
It commits only after package serialisation succeeds. Failed operations preserve
prior edits and handle identity. Successful definition creation preserves existing
paragraph, span, table and cell handles because document XML does not change.
Later paragraph style selection or text edits follow their own invalidation rules.

## Limits

No style rewriting/deletion, character/table/numbering style creation, default
selection, linked/next styles, theme resolution or computed inheritance. Native
save/reopen checks cover registry creation and selection; independent Word
rendering and general schema compatibility remain unverified. The shared
specification defines expected behaviour in `workflows/docx/style-authoring.feature`.
