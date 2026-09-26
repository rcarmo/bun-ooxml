# Cross-language scope limits

Shared workflow outcomes cover the operations named by each scenario. They do
not make Python, Go and Bun API surfaces equivalent. Keep these gaps explicit
when mapping upstream code or adding features.

## Word revisions and patch options

The Python review reports that `word_accept_all_changes` traverses main document
XML only. Separate headers, footers and footnotes, move revisions and formatting
revisions need their own tests. The `office_patch` `track_changes` dispatch argument
was reported unused. These observations are requirements to investigate in a full
port, not compatibility behaviours to reproduce silently.

Bun's current workflow performs direct text replacements in its supported body and
table paragraphs. It has no track-changes switch, accept/reject engine or multi-story
review guarantee. The full port requires per-story traversal and accept/reject
outcomes before those APIs can be marked mapped. Do not imply that a similarly
named Python operation supplies those guarantees.

## Spreadsheet derived values

Bun clears cached `<v>` contents on worksheet cells containing `<f>`. Array and
data-table results can occupy ordinary cells without `<f>`; their presence now
refuses value edits before any model or filesystem commit. This bounded guard
will need replacement by reviewed range-aware invalidation for full parity.

Ordinary formula cache invalidation neither recalculates formulas nor refreshes
chart caches, external-link caches, calculation chains or other opaque derived
parts. These are preserved unchanged. `recalculation-required` reports work that
an engine still needs to do, not proof that every representation is current.

The shared cache fixture deliberately has ordinary formulas, two sheets and no
calcChain. Passing it does not close broader cache/range/external-reference tests.
See `features/xlsx/cache-boundaries.feature` for executable boundaries and
`features/planned/cross-language-followups.feature` for unresolved obligations.

## Extended-comment content type

Pinned source constants disagree:

* `references/fixtures-ooxml/reference-assets/docx/src/docx/commentops.py` and the pinned Python MCP
  `tools/word_tools.py` use
  `application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml`.
* The Go reference at `43eda6e`, `pkg/packaging/constants.go`, uses
  `application/vnd.ms-word.commentsExtended+xml`.

Record this as an unresolved schema/Office compatibility question. A source
majority or one library reopening its own output does not resolve it. Bun has
not implemented this part; select its content type only with an authoritative
schema/reference and independently checked producer/consumer fixture.

## Transport evidence

Python's reported 19-case acceptance is direct-server-call evidence. Its separate
clean-wheel tests exercise stdio. Legacy SSE/TCP availability does not prove
mutation outcomes over those transports; HTTP/SSE/TCP remains unverified here.
Bun executes the shared workflows as native library calls and exposes no MCP
transport in this repository. Neither runner inherits the other's transport proof.
