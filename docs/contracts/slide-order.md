# Existing PPTX slide order

`Presentation.reorderSlides(order)` accepts an exact zero-based permutation of
all current slides and returns `{changed: 0 | 1}`. For three slides, `[2,0,1]`
moves the last slide first. Missing, repeated, fractional, out-of-range or accessor
indexes refuse. Empty decks accept `[]` as a byte-identical no-op.

Only `p:sldIdLst` entry positions change. Each entry retains its exact ID,
relationship ID, attributes and namespace bindings; whitespace between list slots
and presentation XML outside the entries stay unchanged. Slide parts, notes,
relationships, masters, layouts, themes and all other members are byte-identical.
Existing custom-show, extension, protection or unrecognised presentation metadata
refuses because this slice does not reconcile additional order-dependent state.

`presentation.slides` returns a detached array. A Slide object stays tied to the
same part, and its `index` reports its current logical position. Reordering does
not bump slide content versions: previously captured text anchors and table/cell
handles remain usable on their original slide. Content edits keep their usual
stale-handle rules. The public `Slide(presentation, partName, index)` constructor
still accepts its third argument; aliases follow the unique part's current index
instead of retaining the initial hint. No slide cloning, deletion or import is
performed.

The operation checks cached presentation/root relationship bytes before using its
handles. Out-of-band changes require reopening. Existing slide IDs and targets
must be unique, internal and correctly typed. No-op requests still check unsafe
metadata. Appending a slide also refuses stale presentation metadata so it cannot
publish an inconsistent handle list after a raw reorder.

The package transaction serialises before publishing the new handle order. Failed
serialisation preserves earlier edits, archive bytes, indexes and snapshots.
UTF-16/BOM and exact slide-entry spelling are covered by native regressions.

Independent Open XML SDK and LibreOffice PDF page/text checks exercise a two-slide
reversal. This does not establish arbitrary deck rendering or compatibility with
custom shows, sections or extension metadata. Microsoft PowerPoint validation and
full slide composition remain open.

Canonical feature: `workflows/pptx/slide-order.feature` in the shared reference.
Native bindings/tests: `tests/acceptance/slide-order.ts`, `tests/unit/pptx-slide-order.test.ts`.
