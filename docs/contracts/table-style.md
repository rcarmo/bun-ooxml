# Direct Word table-style references

`Table.styleId` reads the direct `w:tblStyle` reference or returns undefined when
absent. `Table.setStyle(id)` writes that reference and returns `{ changed: 0 | 1 }`.
`setStyle(null)` removes it.

```ts
const table = document.addTable(2, 2);
table.setStyle('TableGrid');
console.log(table.styleId); // TableGrid
```

This API stores a reference only. It neither creates a styles part nor requires
that the ID resolve to an existing table style. ECMA-376 Part 1 §17.4.62 specifies
that unresolved references apply no table style. Consequently, writing `TableGrid`
into a document without that definition does not promise a visible grid. This
policy differs from paragraph-style assignment, which requires an existing
paragraph-style definition. Use template documents when a known table-style
appearance is required; effective style resolution is not implemented here.

IDs must be nonempty, well-formed XML strings of at most 255 UTF-16 code units,
without control characters or leading/trailing whitespace. Internal spaces and
Unicode are retained. Empty strings refuse; use null for removal.

## Supported tables and preservation

The target must be a supported top-level table with an explicit grid, 1–100 rows
and 1–100 grid columns. Every row must pass the existing plain row/cell checks;
merged, nested, revised or field-bearing content refuses. These checks do not
validate every grid width or interpret all unrelated property values. The direct
style metadata must be a unique qualified leaf in ordered `tblPr` properties.
Unknown, decorated, duplicate, misordered or lexical metadata refuses even on
same-value or removal requests.

The style reference is inserted first in `tblPr`; missing or self-closing table
properties are supported. New markup declares its Word namespace explicitly.
Existing rows, other table properties, the style registry and unrelated package
parts retain their bytes. Tests cover aliased namespaces, UTF-8 BOMs and UTF-16
save/reopen.

Same-value patches preserve exact archive bytes and handles. Changed patches
expire paragraph, span and table-cell snapshots, but retain table handles because
the grid is unchanged; reacquire cells. Protected writes and stale or externally
changed main XML refuse. Protected direct reads remain available. The prospective
XML must parse and retain collected paragraph text, then package write and
serialization run inside rollback before handles update. The XML editor's
complete output limit also applies.

## Specification and checks

ECMA-376 Part 1, fifth edition, §17.4.62 defines `tblStyle`, and the `CT_TblPrBase`
schema places it first. The [complete specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md)
contains the source PDF. The shared case compares absent and assigned IDs in
memory. Native tests add disk/byte reopen, lexical no-ops, preservation, stale
handles and rollback. An authored unresolved-reference sample and the permanent
Office smoke input pass SDK validation. Neither validation nor getter readback
establishes rendered borders, style inheritance or Office built-in style synthesis.
