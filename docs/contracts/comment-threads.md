# Existing Word comment threads

`inspectCommentThreads(pkg)` groups existing comments into immutable root records
and flat descendant arrays. Roots and replies retain their order in the comments
part; replies keep their direct parent IDs, including nested chains. A reply may
appear before its root in the source. Returned snapshots are detached and deeply
frozen. An empty document returns empty `threads` and `unsupported` arrays.

Like [`inspectComments`](comments-api.md), this reader returns explicit unsupported
findings for unrecognised bodies or unlinked metadata. Malformed graphs throw.
A caller requiring fully supported content must check `unsupported.length === 0`.
Inspection is read-only and does not repair or create any part.

## Complete-thread resolution

`setCommentThreadResolved(pkg, id, resolved)` finds the selected member's root and
changes every member in that existing thread. It returns
`{rootId, commentIds, changed, changedParts}`. IDs in the receipt retain comments
source order; `changed` counts only flags whose state changed. Other threads,
comment bodies, anchors, relationships and unrelated parts are untouched.

Every selected member must already have a valid paragraph ID and linked
`commentEx` entry, even when the requested state is unchanged. Unsupported content
anywhere in the comment package refuses mutation. Protection and external settings
checks precede no-op handling. Decimal IDs accept leading zeros; unsafe or unknown
IDs and nonboolean states refuse. Thread operations support at most 10,000 comments.

All changed `done` values are planned against one XML snapshot. The writer
applies a single edit set, preserves unrelated spelling, prefixes, quotes and
encoding, and validates the entire comment result plus serialization inside the
package transaction. Write/serialization failure restores all flags and retains
prior package edits. A no-op preserves the complete original archive. UTF-8 BOM
and UTF-16 byte order are preserved; the same BOM fix applies to single-comment
resolution.

`setCommentResolved` still edits exactly one comment and never cascades. The thread
method changes both root and replies; it is a different policy from an authored
API that creates missing extension metadata or resolves only the root selected
through a reply. It does not create comments, replies, filters or extensions.

## Evidence and limits

Fourteen native tests use the pinned threaded Word fixture and controlled native
mutations for nested descendants, source-order permutations, exact member/lexical
custody, mixed states, no-ops, missing metadata, protection, unsupported bodies,
rollback, path save/reopen and BOM/UTF-16 encodings. They are unmapped pending
central contracts for this existing-extension, complete-thread policy. The nine
shared authored-comment profiles remain planned; no API-parity credit follows.
No independent Word thread UI or rendering validation has been performed.
