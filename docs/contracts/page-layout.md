# Final DOCX section page geometry

`Document.getPageLayout()` reads the existing final body section's direct page
size, orientation and margins. `setPageLayout(layout)` accepts the complete object
and returns `{changed: 0 | 1}`. Geometry uses integer twips: width/height are
positive; margins, header/footer distances and gutter are nonnegative. This slice
bounds each value to 31,680 twips and requires a positive content rectangle.
Orientation must agree with width/height; square pages allow either orientation.

Only the direct final body `sectPr` is selected. Existing `pgSz` and `pgMar` are
required; missing or ambiguous geometry is refused. Earlier paragraph sections
remain untouched. The writer changes only the selected geometry attributes,
preserving header/footer references, other section metadata and package payloads.
A semantic no-op preserves exact archive bytes and handles. Returned geometry is
a detached value without computed defaults.

Real changes invalidate paragraph/span/cell handles because their XML offsets can
move. Table handles remain valid; reacquire their cells. Stale document XML,
protection, invalid metadata, section revisions, unknown/misordered section
children, lexical barriers, mirrored/book-fold/gutter settings and section bidi
flags refuse atomically. Getter accessors and unknown request fields refuse.
Negative top/bottom margins, implicit sizes and printer-specific layout are outside
this slice. Existing paper-size codes are preserved, not recalculated.

Serialization validates inside the rollback boundary. Native tests cover no-ops,
refusals, namespace aliases/collisions, multi-section custody, UTF-16/BOM,
header relationships, stale handles and disk save/reopen. The independent authored
sample validates with Open XML SDK 3.5.1 and exports through LibreOffice 24.2.7 to
a 792×612-point landscape PDF. This single sample does not establish general
pagination, margin rendering, headers/footers, printer behaviour or Microsoft Word
fidelity.

Canonical feature: `workflows/docx/page-layout.feature` in the shared reference.
Bindings/tests: `tests/acceptance/page-layout.ts`, `tests/unit/docx-page-layout.test.ts`.
