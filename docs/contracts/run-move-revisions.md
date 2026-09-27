# Paired run-move revisions

`inspectRevisions` and `resolveRevisions` accept the opt-in
`profile: 'text-properties-and-moves'`. It adds bounded same-story run moves to
the existing text and run-property operations. Both earlier profiles still
report move markup as unsupported and refuse to resolve it.

```ts
const profile = 'text-properties-and-moves' as const;
const report = inspectRevisions(pkg, { profile });
resolveRevisions(pkg, 'accept', { profile, parts: ['word/header1.xml'] });
await pkg.save('accepted.docx');
```

Use the standalone `OpcPackage` API and reopen any `Document` handles after edits.
The shared [ECMA-376 Part 1](../../references/fixtures-ooxml/specs/ecma-376/index.json)
§§17.13.5.22–.24 and .27–.28 describe inline move wrappers and range markers.
Annex A.1 defines `CT_MoveBookmark`, `CT_MarkupRange` and `CT_RunTrackChange`.
Paragraph-mark move profiles are separate and unsupported here.

## Pairing and admission

A source range and destination range share an exact, nonempty `w:name`. Each side
has its own start/end ID pair. Source and destination IDs need not be equal.
The profile requires one source range and one destination range per name within
one story part. Either side may appear first in source order.

Each range must be three consecutive element children of one paragraph: range
start, one matching `moveFrom` or `moveTo` wrapper, then range end. Only XML
whitespace may occur between them. Different sides may occupy different
paragraphs in the same story. Nested, overlapping, cross-paragraph and cross-story
ranges refuse, along with incomplete pairs or unowned move markup.

Start markers require qualified `w:id`, `w:name`, `w:author` and `w:date`.
Wrappers require qualified ID and author; their date is optional. Supplied dates
use a calendar-valid UTC `YYYY-MM-DDTHH:mm:ss[.fff]Z` spelling with zero to three
fractional digits. Other valid XML Schema date spellings are outside this bounded
policy. Dates are not rewritten. Ends contain only their qualified ID, apart from
namespace declarations. Unknown marker/wrapper attributes refuse.

The safe-integer decimal ID policy is stricter than the general schema: range
starts and wrappers must be unique within the part and must not collide with
supported text or run-property revision IDs. Matching ends reuse their start ID.
Leading-zero decimal spellings are compared numerically.

## Moved content

Wrappers contain only plain `w:r` children, each with one or more ordinary `w:t`
leaves and an optional first `w:rPr`. Moved-from content uses `w:t`, not `delText`.
Text leaves may carry `xml:space`; runs may carry namespace declarations only.
Mixed content, fields, drawings, nested revisions and unsupported properties
refuse. Direct properties use the same bounded set as
[run-property revisions](run-property-revisions.md).

Source and destination must have identical run segmentation, text, `xml:space`
values and direct property trees. Property element and attribute names compare
by namespace URI and local name, ignoring namespace prefix spelling and attribute
order. Scalar values are compared exactly: explicit false and absent flags, or
other semantically similar lexical forms, are not normalised. Moved text must be
nonempty. This prevents resolution from discarding a differently edited version
of the content; combined move-and-edit comparison is unsupported.

Inspection returns two revision records per pair: `move-from` and `move-to`,
with wrapper IDs, author, optional date and concatenated text. Records follow
story and XML source order. Invalid move content produces explicit `move`
findings. A malformed pair blocks all mutation in the selected scope.

## Resolution and preservation

Accept removes the source range's content and unwraps the destination wrapper.
Reject unwraps source content and removes destination content. Both operations
remove all four range markers. Receipts count the two move wrappers, plus other
supported revisions, and list changed story parts in discovery order.

Selected stories preflight together before edits. Unsupported unselected stories
remain byte-identical. The extended profile checks linked protection and external
settings even when no revisions remain; an explicitly empty selection remains an
empty operation. Unknown selections and profile values refuse.

Edits use original XML spans. Surviving runs retain wrapper-local namespace
bindings; ambiguous lifts refuse before mutation. UTF-8 BOM and UTF-16 byte order
survive save/reopen. Changed stories are parsed and rescanned, then the package
is serialized inside one transaction. Reached write, validation or serialization
failures restore every earlier byte, including user edits made before resolution.

## Evidence and limits

Native tests combine moves with insertions, deletions and run-property snapshots
in seven reachable story parts. They compare exact saved output without invoking
the resolver to generate expected text, preserve source and unrelated-member
bytes, and test malformed pairs, selection, encodings, dates and rollback.

The independent Open XML SDK checks one body pair source and accepted/rejected
outputs. A missing-name source must fail schema validation. This does not test
Word's UI, independent move resolution, or multistory rendering.

Shared v0.33 has no cases for this new profile; native declarations are unmapped.
The existing shared profiles and broad multistory obligation are unchanged.
Paragraph, section, table and more general move revisions remain unsupported.
