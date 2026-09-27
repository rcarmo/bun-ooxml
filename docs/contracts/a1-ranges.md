# Direct A1 range parsing

`parseA1Range(source)` parses one direct cell, rectangular range, whole-row range
or whole-column range. It returns a decoded sheet qualifier, an `axis`, and
independent `first` and `last` coordinates with row/column absolute flags.

```ts
import { parseA1Range } from 'bun-ooxml';

const range = parseA1Range("'O''Brien'!$B2:C$4");
// sheet: O'Brien; axis: cell
// first: row 2, column 2, absolute column
// last: row 4, column 3, absolute row
```

Rows are 1–1048576 and columns are A–XFD (1–16384). Letter case is ignored.
Reversed endpoints retain their input order: `A3:B1` starts at row 3 and ends at
row 1. A single cell has two equal but detached coordinates. Whole columns have
row zero; whole rows have column zero. Those zeros mean an absent axis, not a
resolved workbook boundary. Callers must supply actual sheet limits before
using a whole-axis range for editing.

A qualifier applies to both endpoints. Quoted names support spaces, Unicode and
doubled apostrophes. Unquoted names start with a letter or underscore and continue
with letters, digits, underscores or periods. External-book and three-dimensional
references, formula expressions, unions, intersections, spills, leading `=`,
mixed axes, standalone row/column tokens, invalid coordinates and control
characters refuse with `xlsx-range-unsupported`.

Inputs are limited to 1 MiB of UTF-8 and valid JavaScript Unicode strings.
Whitespace outside quoted sheet names is refused, while Go's reference parser
lexes it away. The parser does not check workbook membership, enforce every Excel
sheet-naming rule, normalise ranges for an edit, analyse formula dependencies,
rewrite formulas or calculate values.

The two selected shared scenarios have 13 cases. Their positive cases check only
sheet, axis and first coordinate; native tests additionally check both endpoints
and absolute flags. [Static formula analysis](formula-analysis.md) separately
extracts cell/rectangle references from a bounded expression grammar.
[String-level insertion remapping](formula-remap.md) separately handles supported
row/column shifts; it does not mutate worksheets.
