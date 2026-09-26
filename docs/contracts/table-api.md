# Rectangular table slice

Native rectangular DOCX/PPTX tables. No merge/split/row/column structural mutation
parity is claimed. Existing merged tables can be inspected only when grid mapping
is unambiguous; unsafe edits refuse before mutation. No formatting reconstruction
outside the selected table/cell.

DOCX: `Document.addTable(rows:number, columns:number): Table`, `Document.tables: Table[]`.
Table `rows:number`, `columns:number`, `cell(row,col): TableCell`. TableCell `text:string`
(get/set), supports direct simple paragraphs/runs only; setter may replace the cell
text payload under an explicit formatting policy documented in code (prefer keep
existing first-run formatting). Empty cells valid. Added table goes before sectPr,
valid tblPr/tblGrid/tr/tc/tcPr/p structure, bounded dimensions, table handles refuse
on structural document edits and cell handles refuse after any mutation. Newlines
normalise CRLF/CR and become paragraphs. The first paragraph/run properties are
retained with their namespace declarations. Inherited table styles/merges stay gaps.

PPTX: `Slide.addTable(rows,columns,{x,y,width,height}): Table` with integer EMU geometry.
`Slide.tables: Table[]`, same rows/columns/cell text API. Native graphicFrame/a:tbl,
unique cNvPr id, grid column widths sum width and row heights sum height, xfrm geometry,
valid table properties and trailing p/r instructions. Native reader table/text APIs
must not corrupt existing shape identities. Bounds max10000 cells, positive rows/cols,
positive widths/heights, finite safe integer geometry. Zero x/y allowed. PPTX table
and cell handles become stale after a slide mutation; reacquire from Slide.tables.

Both: zero-based indices, typed range errors, unchanged saved bytes on refusal;
XML escaping and boundary whitespace; saved/reopened table content and opaque
parts verified. Avoid retained mutable XML nodes: reparse after edits and validate
fingerprints or revision counters on handles. Format root source files are disjoint
for parallel work; tests/features named tables-*; parent handles exports/docs.
