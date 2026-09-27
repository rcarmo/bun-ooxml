# Replace plain paragraph text

`Paragraph.setText(text)` replaces the complete text of a supported paragraph and
returns its current handle. A changed edit returns a fresh handle; an identical
text value returns the existing handle after validation.

```ts
let paragraph = document.addParagraph('Old text', { bold: true });
paragraph = paragraph.appendRun(' in another run', { italic: true });
paragraph = paragraph.setText('New text');
// The first run contains New text and remains bold.
// The second run and its italic property remain, with empty text.
```

The new text goes into the first existing `w:t`; all later text leaves are emptied.
Every existing run, paragraph property and run property remains in place. When
there is no text leaf, the editor adds one to the first run. When there is no run,
it creates a plain run. Empty replacement text clears existing text without
removing its runs. Replacing an already empty paragraph is a no-op.

This policy preserves formatting markup, not the appearance of the previous
character ranges. Replacement text uses the first existing text leaf's run
formatting. A leading empty run without a text leaf does not receive text if a
later run already has one. Use [appendRun](append-run.md) to add an independently
formatted run; this setter does not distribute new characters across old ranges.

## Validation and preservation

Text must be a string no longer than 1 Mi UTF-16 code units. Tabs, carriage returns,
line feeds and invalid XML characters refuse. XML metacharacters are escaped;
leading or trailing spaces use `xml:space="preserve"`. Existing namespace
attributes and quotes remain intact when updating that attribute.

Only plain direct runs are supported. Fields, hyperlinks, revisions, unsupported
or misordered properties, lexical barriers and unknown text attributes refuse.
Comments, processing instructions and CDATA inside a text leaf are barriers,
including on same-text calls. The only supported non-namespace text attribute is
`xml:space` with `default` or `preserve`. These are editor restrictions, not a
complete list of valid WordprocessingML constructs.

The held paragraph must match the current document version, identity and exact
main-part XML. Document protection is checked before edits and no-ops. Same-text
calls preserve archive bytes, entities, run boundaries and held handles; they do
not normalise existing whitespace markup. Changed calls expire paragraph, span
and table-cell handles. Table handles remain usable because dimensions are
unchanged; reacquire their cells.

Source outside the targeted text leaves or new run/text insertion remains
unchanged. Tests compare paragraph/run properties, a sibling paragraph and
unrelated package members, and check UTF-16 byte reopen. The prospective document
must parse, retain all other paragraph texts and retain table dimensions before
package write and serialization. The transaction rolls back failures before
updating model handles. Complete XML output is also subject to the XML editor's
8 Mi UTF-16-unit limit; this is not a peak-memory guarantee.

ECMA-376 Part 1, fifth edition, §17.3.2.25 defines runs and §17.3.3.31 defines text
and whitespace handling. The [complete specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md)
contains the source PDFs. The shared text-getter cases check five in-memory
values; separate native tests cover disk/byte reopen and refusals. An authored
sample passes Open XML SDK validation, and the Office smoke test renders replaced
text. General formatting, pagination and visual equivalence have not been tested.
