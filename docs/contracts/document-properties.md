# Direct Word title-page and background properties

`Document.getDocumentProperties()` returns `{ titlePage, backgroundColor }` for a
single-section document. `setDocumentProperties(patch)` updates supplied values
and returns `{ changed: number }`, counting changed fields.

```ts
document.setDocumentProperties({
  titlePage: true,
  backgroundColor: 'EEEEEE',
});
```

`titlePage` is Boolean or null. True/false writes an explicit `w:titlePg` value;
null removes it. Absence reads as null, distinct from explicit false. Word treats
an absent marker as disabled. The marker enables separate first-page header/footer
selection; it does not create those headers/footers or insert a title page.

`backgroundColor` is a six-digit RGB string, retaining case, `auto`, or null.
Null removes the plain `w:background` element. An existing empty background reads
as null; removing that empty element still counts as a change. This API does not
resolve an `auto` colour. Undefined fields are left alone; an empty patch refuses.

## Supported documents

The document must contain exactly one `w:sectPr`, directly at the end of the sole
body. Documents with earlier section breaks refuse, even for background-only
patches. A self-closing final section is supported without inventing page geometry.
This single-section limit makes the shared first-section example unambiguous.
The existing final-section page-geometry API retains its separate scope.

Only a plain direct background colour is supported. Themed, shaded or drawing
backgrounds, decorated leaves and invalid values refuse, including removal or
same-value requests. Root content must be an optional background followed by one
body. Section metadata must have supported ordering and no revisions or lexical
barriers. Untouched section metadata receives structural checks, not complete
value interpretation or schema validation. Existing header/footer references and
body content remain in place.

## Preservation and failures

Same-value patches preserve exact archive bytes and handles after validation.
Changes replace only selected property elements or insert missing ones at their
schema positions. Authored elements bind their Word namespace explicitly, so
aliases and conflicting inherited prefixes are safe. Native tests compare exact
existing main-part fragments, unrelated package parts and UTF-16 save/reopen.
The shared property-commit path retains an existing UTF-8 BOM as well.

A changed patch expires paragraph, span and cell handles; table handles remain
usable. Protected writes refuse, including no-ops. Reads do not mutate protected
documents. Both operations reject externally changed main-part XML instead of
using stale snapshots. The prospective XML must parse and retain every collected
paragraph text; package write and serialization are transactional before model
handles change. The XML editor's 8 Mi UTF-16-unit complete output limit applies.

## Specification and checks

ECMA-376 Part 1, fifth edition, §17.2.1 defines background markup and §17.10.6 defines
`titlePg`. See the [complete specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).
When the first section enables `titlePg` but has no first-page header/footer,
those areas are blank rather than borrowing another header/footer type.

The shared case checks two in-memory getter values. Native tests separately cover
path save/reopen, exact preservation, aliases, stale handles and rollback. An
authored sample and the permanent Office smoke input pass SDK validation. PDF
smoke assertions check text/page counts, not background colour or first-page
header rendering. Pagination, theme resolution and multi-section editing remain
outside this API.
