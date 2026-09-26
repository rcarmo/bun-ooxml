# Editing existing speaker notes

`slide.inspectNotesText()` returns a frozen `NotesAnchor` for the existing body
placeholder. `slide.replaceNotesAt(anchor, text)` replaces that body through a
fresh target and returns `{ changedParts }`. It never creates a notes part or
changes another placeholder.

```ts
const deck = await Presentation.open('slides.pptx');
const slide = deck.slides[0]!;
const target = slide.inspectNotesText();
slide.replaceNotesAt(target, 'First point\n\nClosing point');
await deck.save('edited.pptx');
const reopened = await Presentation.open('edited.pptx');
console.log(reopened.slides[0]!.inspectNotesText().text);
```

The edit requires one internal notes relationship, one incoming owner, the notes
content type and exactly one direct body placeholder. Presentation protection,
text locks, fields, breaks, extensions, unknown namespaces and unsupported
paragraph/run properties refuse. The older `readNotesText()` API remains a
separate broad read-only inspection path.

Inputs are JavaScript strings with LF line separators. Tabs, CR, control
characters and invalid Unicode refuse. A single text leaf uses an exact lexical
text splice where possible. Structural replacement preserves the first
paragraph's supported properties, the first run's properties and the first
paragraph's ending properties. It never borrows later-run bold formatting.
Leading and trailing newlines create empty paragraphs, and text leaves receive
`xml:space="preserve"` when required.

The template subset includes paragraph alignment/margins, a character bullet and
default run properties; direct run bold/italic/font-size/language and selected
attributes; solid RGB colour and a Latin font. Text-body properties and list-style
containers must be empty in this first profile. Unknown or conflicting property
structure refuses. This is a bounded editor policy, not full DrawingML support.

The target captures notes, slide and relationship bytes. Foreign or forged targets,
changed notes or relationship graphs, and targets held across a successful notes
change refuse. Reacquire a target after editing. An identical-text operation keeps
the original archive bytes and target. Failed validation or serialization rolls
back the attempted mutation; prior valid state and target remain usable.

Unrelated parts and non-body placeholders retain their bytes. The notes part's
original encoding is preserved. Tests save and reopen an edited fixture and
compare all original member payloads. Multiline template cases separately inspect
in-memory text and selected property fragments; no rendering or complete schema
claim is made.

Seven existing shared notes cases are bound. Their historical IDs and step text
name Go, but the Bun bindings execute the same pinned inputs and predicates.
The foreign/invalid-byte compound case remains planned: a JavaScript string API
cannot exercise the raw invalid UTF-8 byte `FF` input. Bun separately tests foreign
handles, NUL, tabs, CR and unpaired surrogates without claiming that missing case.
