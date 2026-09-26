# Existing paragraph style assignment

`Paragraph.styleId` reads a direct `pStyle` ID without computing inheritance.
`Paragraph.setStyle(styleId: string | null)` assigns an existing paragraph style
or removes only that override. It returns `{ changed: 0 | 1 }` for the paragraph.

```ts
const paragraph = document.paragraphs[0]!;
paragraph.setStyle("Heading1");
// Reacquire after a real change.
document.paragraphs[0]!.setStyle(null);
```

A non-null ID requires one exact internal styles relationship, the Word styles
content type, a valid root and exactly one direct matching paragraph-style
entry. Missing or character styles, duplicate selected IDs and ambiguous/external
relationships refuse. `Document.addParagraph(..., {style})` uses the same strict
lookup. Style definitions, links and all unrelated package members are unchanged.

Null removes only `pStyle`; it leaves the `pPr` container and other properties
alone and requires no registry. The getter can report an unknown stored ID;
it neither validates definitions nor chooses a default. Malformed direct metadata
still refuses. No operation creates styles or evaluates effective formatting.

## Editing and preservation

Assignment preserves paragraph text, direct run formatting and unrelated
paragraph properties. New `pStyle` is inserted first in `pPr`; existing style
markup can be replaced by a namespace-qualified element. Existing alias/default
namespaces and supported UTF-16/BOM encoding survive save/reopen. A semantic no-op
preserves exact archive bytes and handles, including different lexical spellings.

Real changes invalidate paragraph, span and table-cell snapshots; reacquire them.
Table handles stay valid because no grid changes occur. Fresh spans can still edit
text after style assignment. Serialization validation runs inside rollback, and
model snapshots update only on success; failure preserves earlier edits/handles.

Protection, external/invalid settings, unsupported text topology, property
revisions, mixed lexical content, duplicate/misplaced/unknown paragraph properties,
wrong namespaces and stale snapshots refuse before mutation. Empty paragraphs and
self-closing property containers can receive a style. Error codes use the existing
`docx-style-*`, `docx-format-*`, stale and unsupported-topology families.

The shared specification defines expected behaviour in
`workflows/docx/paragraph-style.feature`. This API does not edit paragraph layout,
create styles or compute inheritance. Independent Office rendering is unverified.
