# Static spreadsheet formula references

`analyzeFormulaReferences(source)` returns static A1 cell or rectangle references
in source order. Each record contains `sheet`, `axis: 'cell'`, `first` and `last`
coordinates with independent absolute flags, plus half-open `start`/`end` UTF-8
byte offsets into the original expression.

```ts
import { analyzeFormulaReferences } from 'bun-ooxml';

const source = `IF(A1="B2",'O''Brien'!$C$4,SUM(D1:E2))`;
const references = analyzeFormulaReferences(source); // three records
const bytes = new TextEncoder().encode(source);
const firstToken = new TextDecoder().decode(
  bytes.slice(references[0]!.start, references[0]!.end),
); // A1
```

String contents do not create references. Quoted sheet names decode doubled
apostrophes, while each span retains the complete original spelling, including
sheet qualifiers and `$` markers. Reversed rectangle endpoints retain their
source order. Unqualified references have an empty sheet name. Returned records
and coordinate objects are detached; modifying one result does not affect later
calls. Offsets are bytes, not JavaScript string indexes.

## Accepted grammar

The bounded grammar accepts an optional leading `=`, numeric literals with decimal
and exponent forms, `TRUE`/`FALSE`, quoted strings with doubled double quotes,
parentheses, unary `+`/`-`, postfix `%`, and binary `+ - * / ^ & = <> < > <= >=`.
ASCII spaces are permitted between tokens. Tabs, line breaks, control characters
and invalid Unicode refuse.

Calls are limited to `IF`, `SUM`, `LOG10`, `ABS`, `MIN`, `MAX`, `AVERAGE`, `COUNT`,
`COUNTA`, `ROUND`, `AND`, `OR` and `NOT`, case-insensitively. Each call requires at
least one nonempty argument; commas separate arguments. Argument count and value
types are not checked against Excel function definitions. A successful parse
therefore does not establish that Excel would evaluate the expression.

Reference tokens reuse the [direct A1 parser](a1-ranges.md) for sheet and grid
bounds. Expression operands accept only cells and rectangles, even though that
separate parser also supports whole axes. Defined names, structured references,
external workbooks, sheet spans, unions/intersections, array literals, error
literals, implicit intersection, spills and dynamic calls such as `INDIRECT` and
`OFFSET` refuse. Unknown functions also refuse. This function does not evaluate,
rewrite references, inspect workbook sheets or update formula caches.

## Refusals and limits

Unsupported syntax throws `OoxmlError` with code `xlsx-formula-unsupported`.
Resource limits use `xlsx-formula-limit`. A late failure throws without returning
a partial reference array. No package or caller-owned object is modified.

Limits are 1 MiB of UTF-8 input, 100,000 tokens, 10,000 references and 128 recursive
expression levels. The parser also checks the string length before allocating
its UTF-8 offset table. These are processing/input limits, not peak-memory
measurements. The function uses Bun/TypeScript only.

Four existing shared scenarios supply 19 cases for reference counts and spans,
quoted-sheet flags, unsupported input and literal punctuation. The accepted
punctuation cases check parse success, without a full reference-value predicate.
Native tests separately compare exact reference slices and a finite 288-expression
matrix. The compound shared matrix remains planned because it also requires
insertion remapping, which this API does not provide. No calculation, remapping,
saved-workbook, rendering or cross-runtime parity credit follows from analysis.

The shared [static formula profile](../../references/fixtures-ooxml/contracts/go-formula-references.md)
defines these API examples. The supported grammar is an editor policy, not a
claim of complete ECMA-376 formula-language conformance.
