# Cross-language limits

The shared specification defines expected outcomes. Each library implements a
different subset; passing selected workflows does not establish general
interchangeability.

Word review supports bounded direct run insertion/deletion, selected linked
stories and exact tracked replacement. `patchOffice` can dispatch one tracked
Word replacement with explicit author/date and separate preview/commit revision
receipts. Existing comment inspection and per-entry done flags are supported;
Python's root-thread redirection and metadata creation remain distinct operations.
General comparison, moves/property/table revisions, comment authoring and
multi-target tracked batches are unsupported.

Spreadsheet edits invalidate supported worksheet formula caches without computing
results. Array/dataTable topologies refuse value edits; chart/external-link caches
and calculation chains are preserved, not refreshed.

For commentsExtended metadata, use the shared specification's
ContentTypeCommentsExtended and ContentTypeCommentsExtendedSpecified facts.
The observed alias is disputed; vendor metadata and the pinned fixture agree on
the specified value. Independent Office validation of comment authoring and
reopening is unavailable.
