# Inspecting existing spreadsheet comments and VML

`inspectWorksheetComments(package, worksheetPart)` reads current metadata from an
`OpcPackage` and returns detached records in document order:

```ts
import { OpcPackage, inspectWorksheetComments } from 'bun-ooxml';

const pkg = await OpcPackage.open('book.xlsx');
const result = inspectWorksheetComments(pkg, 'xl/worksheets/sheet1.xml');
// result.comments: { reference, authorId, author, text }[]
// result.commentsRelationship / result.vmlRelationship: { id, type, part } | null
```

Pass the actual worksheet part resolved from the workbook relationship graph;
`sheet1.xml` is an example, not a filename-discovery rule. Each call reads live
package bytes. Mutating the result does not change the package or future results.
Neither protection nor a previous failed inspection changes this read-only policy.

Existing worksheet comments can use two separate relationships: one to the
comment XML and another to a VML drawing. The worksheet's `legacyDrawing` element
identifies the VML relationship through an Office-relationships namespace `id`
attribute. Resolve that ID in the worksheet's relationship part; a matching local
attribute name in another namespace is insufficient.

The comment relationship has type
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments`.
Its XML stores cell references and text. The VML relationship instead has type
`http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing`;
it describes the legacy drawing and must not be substituted for comment text.

The shared [existing-graph case](../../references/fixtures-ooxml/workflows/xlsx/comment-vml-custody.feature)
uses a pinned workbook. Bun's OPC/XML readers resolve `anysvml` to
`xl/drawings/commentsDrawing1.vml`, resolve the separate comment relationship to
`xl/comments/comment1.xml`, and compare the exact A2 and A3 comment strings.
The binding also requires unchanged input and archive bytes. Negative controls
alter namespace URIs, IDs, types, targets, comment values, duplicate records and
payloads. An aliased-namespace example checks rich comment text assembled from
SpreadsheetML text leaves.

The native inspector requires a Transitional worksheet, the expected worksheet,
comments and VML content types, and unambiguous internal relationships. Missing
dependencies, external targets, fragment/query targets, duplicate legacy markers
and mismatched namespaces refuse. Comment and VML parts are independently optional:
no relationships returns an empty list and two null references; a VML-only graph
returns no comments. A VML relationship must match a direct `legacyDrawing` marker.

The comment subset is ordered `authors` then `commentList`, with at most 10,000
authors and 10,000 comments. References must be unique uppercase A1 cell addresses
within Excel bounds, without dollar signs or leading zeroes. Author IDs must be
canonical zero-based decimal indices into the authors array. This is an API
admission policy, not a complete schema validator. Optional `shapeId` values are
retained metadata, not validated or returned. Inspection does not search for
orphan parts or count owners across other worksheets.

Text can be empty, a plain SpreadsheetML `t` leaf, or rich runs with one `t` and
optional first `rPr`. Supported run-property elements are inspected structurally
but their visual values are not interpreted. Unknown text structures, phonetic
annotations, extensions, nested text, and lexical markup inside comment content
refuse rather than produce partial text. XML parser size/depth/node limits also
apply. Original encoding and bytes are untouched. Module policy refusals use
`xlsx-comments-unsupported`; lower-level OPC/XML failures retain their own codes.

The canonical binding calls this inspector and compares the returned values.
It adds no worksheet-comment editing API, VML shape validation or rendered-note
guarantee. VML payloads remain opaque, even if their bytes are not valid XML.
The five shared row-editor/numeric-editor policy cases are not executed. Their
admission policies differ, and no behaviour is inferred from this inspection.

ECMA-376 Part 2 §6.5.3.4 defines relationship Type and Target. Part 1 informative
Annex L.2.6.3 illustrates separate comment and VML data; informative Annex L.5.1
deprecates VML for new drawings. Deprecation does not permit discarding retained
package data. The complete documents are in the
[shared specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).
