# Direct-body paragraph anchors

`Document.inspectBodyAnchors(query?)` returns a frozen array of frozen paragraph
anchors in body order. Each anchor contains `kind: 'docx-body-anchor'`, `type`,
`bodyIndex`, `text` and `outlineLevel`. Top-level tables count toward `bodyIndex`,
but their cell paragraphs are excluded. Final section properties do not count.

```ts
const document = await Document.open('source.docx');
const anchor = document.inspectBodyAnchors('delivery')
  .find(anchor => anchor.text === 'Delivery approach');
if (!anchor) throw new Error('Expected delivery paragraph');
document.insertParagraphAfter(anchor, 'New delivery detail');
await document.save('edited.docx');
```

Classification uses only the paragraph's direct `w:outlineLvl`: values 0–8 produce
`section_heading`; an absent value or explicit 9 produces `paragraph`. A style
named `Heading1` alone is insufficient. Inherited outlines, style-name heuristics,
rendered headings and section boundaries are not resolved. See
[direct outline levels](paragraph-properties.md) for ECMA-376 §17.3.1.20 semantics.

Queries use JavaScript lowercase substring comparison, without locale rules or
Unicode normalisation. Omitted or empty queries return all candidates; a missing
match returns an empty array. Other query types and strings over 4096 UTF-16 code
units refuse. The entire direct body and every candidate paragraph is checked
before filtering. Unsupported fields, revisions, malformed selected properties,
unknown direct body children or lexical body barriers refuse rather than yield a
partial list. Existing XML size/depth/node limits bound discovery. This API omits
table-cell paragraphs, headers, footers and other stories by definition.

`Document.insertParagraphAfter(anchor, text, options?)` returns the inserted
`Paragraph`. It accepts only an anchor issued by that document, at the same
mutation version and with exact main-part bytes still present. Copied, forged,
foreign and stale anchors refuse with `docx-stale-anchor`. A BOM-only external
change also invalidates a captured anchor. Anchor records retain the captured
main-part buffer while held; anchors from one discovery call share that buffer.

Insertion delegates to [indexed body insertion](body-insertion.md), including
plain text/options, protection, output bounds, preservation and rollback checks.
Successful insertion expires prior paragraph, cell, table and anchor handles.
Validated no-ops in existing properties and refused writes keep current anchors
usable. Discovery is read-only and is allowed on protected documents; insertion
still refuses. No save occurs until the caller requests it. UTF-8 BOM and UTF-16
member encoding survive insertion and save/reopen.

Three shared [anchor-discovery cases](../../references/fixtures-ooxml/workflows/docx/anchor-discovery.feature)
run through this API: listing headings/paragraphs, case-insensitive filtering and
saved insertion after a selected heading. Their authored documents use explicit
direct outline levels and are saved/reopened before discovery. No Python or MCP
transport runs in these bindings. Document-map counts and discovery-tool hints
remain planned; these anchors do not provide those operations.
