# Word tracking preference

`Document.enableTracking(author)` saves an enabled `w:trackRevisions` setting and
sets the session author. `disableTracking()` saves a disabled setting and retains
the session author. Both return `{changed, changedParts}`. An unchanged setting
returns zero and preserves the original archive bytes.

```ts
const document = await Document.open('draft.docx');
document.enableTracking('Reviewer');
document.setTrackAuthor('Second reviewer');
await document.save('review.docx');
```

`trackChangesEnabled` reads the saved setting. `trackAuthor` and
`setTrackAuthor(author)` hold session metadata; reopening starts with an empty
author. Authors must be nonblank, XML-safe strings. No author is written into the
settings part.

These methods do not wrap Bun text edits in revision markup. Use
[`trackedReplace`](review-api.md) or the explicit [`patchOffice`](workflow-api.md)
`trackChanges` option for tracked replacements. Enabling the saved preference
requests tracking by an application that honours it; automatic tracking during
ordinary Bun mutations is outside this API.

## Saved settings

ECMA-376 Part 1 §17.15.1.89 defines `trackRevisions`; its absence means disabled.
The `CT_Settings` sequence determines its position among neighbouring settings.
The writer accepts Transitional settings with known, correctly ordered direct
children. It preserves siblings, XML encoding and unrelated package members.
An existing setting accepts absent `val`, `1`, `true`, `on`, `0`, `false` and `off`.
Disabling an existing enabled setting writes explicit `0`.

When no settings part exists, enabling creates `word/settings.xml`, its content
type and one internal relationship. Disabling an absent setting creates nothing.
The writer refuses occupied names, orphan or shared settings, external or duplicate
relationships, unexpected roots/content types, duplicate or misplaced settings,
unknown extension children, decorated roots and malformed on/off values. It does
not repair or adopt an ambiguous graph. Enforced or indeterminate document
protection refuses mutations, including no-ops.

Writes and serialization run inside the package transaction. A failure leaves
both package members and the session author unchanged. Main-document handles
stay valid because settings changes leave the document body untouched.

## Checks and limits

The shared `@id-docx-go-track-author-toggle` case checks the getter sequence only.
`tests/unit/docx-tracking-settings.test.ts` adds path save/reopen, custom-part and
namespace handling, UTF-16 retention, byte custody, refusal and injected-fault
checks. `make office-oracles` validates enabled and disabled output packages with
the Open XML SDK. The SDK check covers schema validity; Office UI behaviour and
automatic redline generation are untested.

Shared v0.28 adds seven [settings contracts](../../references/fixtures-ooxml/contracts/tracking-settings.md)
with 24 cases. `tests/acceptance/tracking-outcomes.ts` saves and reopens outputs,
checks exact sibling and unrelated bytes, retains UTF-8 BOM and UTF-16 encodings,
and verifies same-state no-ops, dual refusals and rollback. The BOM case exposed
and fixed a writer bug that dropped the UTF-8 marker.

`docs/behaviors/tracking-settings-mappings.json` records fourteen declarations:
twelve native tests and two shared-outcome wrappers. The getter wrapper selects
one case and the settings wrapper selects 24; other associations are scenario-only.
Every mapping retains explicit gaps and partial status. Python and Go receive no
execution credit from these Bun results.
