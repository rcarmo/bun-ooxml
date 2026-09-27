# Inspecting existing spreadsheet comments and VML

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

This is read-only inspection through the existing OPC/XML APIs. It adds no
worksheet-comment editing API, VML shape validation or rendered-note guarantee.
The five shared row-editor/numeric-editor policy cases are not executed. Their
admission policies differ, and no behaviour is inferred from this inspection.

ECMA-376 Part 2 §6.5.3.4 defines relationship Type and Target. Part 1 informative
Annex L.2.6.3 illustrates separate comment and VML data; informative Annex L.5.1
deprecates VML for new drawings. Deprecation does not permit discarding retained
package data. The complete documents are in the
[shared specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).
