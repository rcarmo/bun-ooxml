# Run-property revision resolution

`inspectRevisions` and `resolveRevisions` accept an optional
`profile: 'text-and-run-properties'`. The default `text-only` profile still
reports `w:rPrChange` as unsupported and refuses to resolve a selected story
containing it.

```ts
const options = { profile: 'text-and-run-properties' as const };
const report = inspectRevisions(pkg, options);
resolveRevisions(pkg, 'reject', { ...options, parts: ['word/header1.xml'] });
await pkg.save('reviewed.docx');
```

These standalone methods operate on `OpcPackage`. Reopen a `Document` after
resolution; existing document handles do not refresh automatically.

## Current and previous properties

ECMA-376 Part 1 §17.13.5.31 defines run-property changes: the enclosing `w:rPr`
holds current properties, while `w:rPrChange/w:rPr` holds the complete previous
set. §17.13.5.30 describes paragraph-mark changes, which this profile refuses.
The [shared specification index](../../references/fixtures-ooxml/specs/ecma-376/index.json)
identifies the normative PDF; Annex A.1 defines `CT_RPrChange` and `CT_RPrOriginal`.

Accept removes only the change element. Reject replaces the entire current
`rPr` with the saved `rPr`, including an empty saved set. Properties missing from
the previous set are removed, not inherited from the current set. Run text is
unchanged. Inspection reports `kind: 'run-properties'`, the part, revision ID,
author, optional date and the affected run's text. Receipts count one revision
per property snapshot plus any supported insertion/deletion wrappers.

## Admission

Only the unique first `rPr` in a direct paragraph text run may contain a change.
The change must be the last property child, have a qualified decimal ID and
qualified author, and contain exactly one saved `rPr`. The existing per-part
safe-integer ID uniqueness policy also applies across text and property changes.

Both current and previous sets must use ordered, nonduplicate direct properties:

- Boolean bold, italic, caps, small caps, strike, double strike, outline, shadow,
  emboss, imprint and hidden flags, including complex-script bold and italic;
- direct colour, underline, highlight and vertical alignment;
- positive integer half-point sizes (`sz` and `szCs`);
- matching direct ASCII/high-ANSI font names.

The profile refuses unknown, themed, foreign, decorated or conflicting properties,
mixed content, nested changes, non-text runs, paragraph-mark revisions and missing
snapshots. Moves, paragraph, section, table and numbering changes are unsupported.
A supported result is a bounded editing policy, not full schema validation.

## Preservation and refusal

The resolver discovers the same linked stories as the text-only profile. Selected
stories are preflighted together; an unsupported later story prevents all edits.
Other stories retain exact member bytes. The extended profile checks linked
settings before a same-state request, refusing enforced protection, malformed
settings roots and external settings relationships. An explicitly empty part
selection remains an empty operation.

Reject preserves the saved element's lexical spelling and transfers namespace
bindings from the removed current `rPr` and change element. A saved element's own
bindings take precedence. Unrelated text and XML are unchanged. UTF-8 BOM and
UTF-16 byte order survive edits; UTF-8 BOM preservation also applies to default
text-only resolution.

All edits use the original story offsets in one package transaction. The extended
profile reparses and rescans changed stories before serialization. A write,
validation or serialization failure restores every member, including prior edits.

## Verification limits

Native tests check both actions across seven reachable story parts, selected
scope, default refusal, malformed snapshots, namespace shadowing, three encodings,
empty previous sets, protection before no-op and reached transaction faults.
The Office oracle validates one source snapshot and its accepted/rejected body
outputs using the independent Open XML SDK; a missing-author control must fail.
No independent Word UI or multistory rendering comparison has been performed.

The [shared opt-in profile](../../references/fixtures-ooxml/contracts/run-property-revisions.md)
adds eight scenarios / 33 cases. Bun builds expected XML independently of the
resolver and checks saved output, exact bytes, receipts and refusal outcomes.
Twelve native tests and two outcome tests have partial mappings with explicit gaps.
The ten earlier text-only cases are unchanged. The broader multistory requirement
still includes unsupported move and property-change forms.
