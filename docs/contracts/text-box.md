# Positioned PPTX text boxes

`Slide.addTextBox(text, geometry, options?)` appends one plain text box and returns
`{shapeId, partName, paragraphCount}`. Geometry is `{x,y,width,height}` in EMUs.
Positions are nonnegative, extents positive, and every value is an integer no
greater than 2,147,483,647. The writer does not constrain boxes to the slide edges.
Options accept a nonempty `name` and boolean `bold`/`italic` direct flags.

CRLF and CR become LF; each line becomes one DrawingML paragraph, including empty
lines. Text is escaped and marked for whitespace preservation. The shape has a
rectangular transform, no fill/line and explicit no-autofit. It has no placeholder
binding. Omitted formatting can still be affected by the presentation's styles;
the API does not calculate the effective appearance.

Only the selected slide XML changes. Existing shapes, extension payloads, layouts,
masters, themes, notes, relationships and content types are preserved. The new
shape binds its own PresentationML and DrawingML namespaces. UTF-16 encoding/BOM
are retained. Shape IDs use the maximum numeric slide-local `cNvPr` ID plus one,
including nested groups. Invalid/duplicate IDs and exhaustion refuse.

Creation requires an unambiguous direct shape tree with leading group properties,
no unsupported top-level children or lexical barriers, and an absent or zero root
group transform. A trailing extension list is preserved after the new shape.
The writer checks shape nonvisual identity containers but does not validate every
existing shape against the full Office schema.

The operation re-reads current slide XML and checks that its handle still resolves
at the same presentation index. Presentation modification protection refuses.
Plain-data options/geometry reject getters and unknown fields. Serialisation runs
inside the package transaction; refusal leaves bytes and slide versions unchanged.
A successful append invalidates same-slide text anchors and table/cell handles.
Handles on other slides remain usable. Reacquire same-slide handles after insertion.

Tests cover saved/reopened text, geometry, flags, names, identity, existing-shape
preservation, real-fixture payload custody, stale handles and injected rollback.
Independent PowerPoint rendering/schema validation, inherited layout formatting,
text fitting, arbitrary shapes and shape deletion/cloning remain open.

Canonical contract: `workflows/pptx/text-box.feature` in the shared reference.
Native tests: `tests/unit/pptx-text-box.test.ts` and `tests/acceptance/text-box.ts`.
