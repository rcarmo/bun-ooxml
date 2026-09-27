# Direct Word table-cell properties

`TableCell.directProperties()` reads selected direct `tcPr` values.
`setProperties(patch)` changes them without replacing cell paragraphs or changing
the table grid.

```ts
const table = document.tables[0]!;
table.cell(0, 0).setProperties({
  widthTwips: 2400,
  verticalAlign: 'center',
  textDirection: 'tbRl',
  shading: 'FFFF00',
  topBorder: { style: 'single', size: 8, color: '000000' },
});
const direct = table.cell(0, 0).directProperties();
```

Supported values:

| Property | Editor policy |
|---|---|
| `widthTwips` | Integer 0–31680; writes an explicit `dxa` preferred width |
| `verticalAlign` | `top`, `center`, `bottom` |
| `textDirection` | `lrTb`, `tbRl`, `btLr` |
| `shading` | Six RGB hex digits or `auto`; clear pattern, automatic foreground |
| `topBorder` | `{ style: 'single', size: 2–96, color: RGB or 'auto' }`; size is in eighth-points |

Null removes the direct property; undefined preserves it. Empty patches refuse.
The getter returns null for absent properties and a detached top-border object.
It does not resolve table styles, border conflict rules or layout defaults.
Changing a cell width leaves the table grid unchanged: a preferred width is not
a guaranteed rendered width.

A real change returns `{ changed: 1 }` and expires held cell and paragraph
snapshots. The table handle remains usable because its grid has not changed;
obtain a fresh cell through it. An identical request returns `{ changed: 0 }`
without changing archive bytes or handles. Protection and topology checks still
run for identical requests.

## Preservation and refusals

The supported cell must contain plain paragraphs in an unambiguous, explicit
table grid. Merged or spanned cells, nested tables, missing paragraphs, property
revisions, unknown or misordered cell properties, duplicate containers and lexical
barriers refuse. Linked document protection and stale source XML also refuse.

The five supported direct properties are decoded before any mutation, even if
only one field is requested. Other permitted `tcPr` children are checked for
structural order and revisions, without interpreting all their metadata. Percentage or automatic width metadata, patterned or
themed shading, and decorated top borders therefore refuse instead of losing
attributes. A top border accepts only explicit single style, size and colour.
Other border fragments are preserved without being interpreted as effective
borders; arbitrary table-property schema validity is not checked.

Selected property leaves can be rewritten with WordprocessingML-qualified names.
Cell paragraphs, sibling cells, the table grid, unselected property fragments and
other package payloads retain their source bytes. Removing the last top border
can leave an empty `tcBorders` container. Edits and serialization checks share a
transaction, so a staged write or serialization failure preserves the original
package and handles. Existing UTF-16 encoding survives path save and reopen.

## Specification and coverage

ECMA-376 Part 1, fifth edition (October 2016), defines cell shading in §17.4.32,
cell borders in §17.4.66, preferred cell width in §17.4.71, text direction in
§17.4.72 and vertical alignment in §17.4.83. The complete PDF is in the
[shared specification index](../../references/fixtures-ooxml/specs/ecma-376/README.md).
The numeric bounds, selected border style and accepted metadata forms above are
editor limits; the standard describes broader forms and layout behaviour.

Two shared scenarios inspect direct values in memory, including the presence of
a top border. Native tests separately check save/reopen, sibling/grid custody,
UTF-16, protection, stale targets and rollback. One authored sample passes Open
XML SDK's Office2019 schema profile. Rendering, preferred-width negotiation,
style inheritance and border conflict resolution remain untested.
