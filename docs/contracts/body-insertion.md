# Insert a Word body paragraph

`Document.insertParagraph(index, text, options?)` creates a plain paragraph at a
zero-based body position and returns its fresh paragraph handle. Each top-level
paragraph or table counts as one position. Cell paragraphs and the final body
section properties do not count.

```ts
const document = Document.create();
document.addParagraph('First');
document.addParagraph('Third');
const paragraph = document.insertParagraph(1, 'Second', { bold: true });
// document.paragraphs now reads First, Second, Third.
```

Index zero inserts before the first block. An index equal to the block count
inserts after the last block, before final `sectPr` if present. Empty text still
inserts a paragraph. Self-closing and sectionless bodies are supported; insertion
does not create section properties. Earlier section properties inside an existing
paragraph stay with that paragraph.

The returned paragraph's `index` uses the document's flattened paragraph order,
including table-cell paragraphs. It can differ from the supplied body index.
Successful insertion expires all previous paragraph, span, table and cell handles.
Reacquire them from the document.

## Inputs and refusals

The index must be an integer between zero and the number of direct body blocks.
Text must be a string of at most 1 Mi UTF-16 code units, without tabs or line
breaks. Invalid XML characters refuse; metacharacters are escaped and boundary
spaces use `xml:space="preserve"`.

Options accept plain data: optional Boolean `bold` and `italic`, and a nonempty
`style` ID that resolves to an existing paragraph style. Accessors, inherited
options, unknown fields and unknown styles refuse. Style definitions are never
created implicitly. These restrictions apply to `insertParagraph`; the existing
`addParagraph` API retains its own contract.

The document must have a WordprocessingML root and exactly one direct body. The
body must contain only paragraphs, tables, whitespace and at most one final
`sectPr`. Unknown body children, comments, processing instructions, mixed text,
duplicate or misplaced final section properties refuse. Existing paragraph and
table subtrees are preserved without interpreting every child; an untouched field
inside an existing paragraph does not prevent inserting a separate body block.

Document protection and exact current main-part XML are checked before editing.
An out-of-band main-part change refuses rather than overwriting that change.

## Preservation and saved checks

The insertion retains all existing body child fragments and unrelated package
members. Newly authored elements declare their Word namespace explicitly, so
aliased bodies and foreign lexical `w` bindings remain safe. UTF-16 encoding is
retained on serialization. The prospective document must parse, contain the new
paragraph and retain every existing collected paragraph text and table dimension.
These dimension checks do not validate all table semantics.

Package write and serialization run inside a transaction; failure retains package
bytes and prior handles. Complete output is subject to the XML editor's 8 Mi
UTF-16-unit limit. This bounds output size, not peak memory.

ECMA-376 Part 1, fifth edition, §17.2.2 defines the document body and §17.3.1.22
defines a paragraph. See the [specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md)
for the source PDFs. The shared body-order case checks in-memory block counts and
First/Second/Third order. Native tests add real path save/reopen, mixed-table
indexes and failure controls. An authored sample passes Open XML SDK validation;
the Office smoke test checks inserted text order in one PDF. General pagination,
section restructuring and imported body content are outside this API.
