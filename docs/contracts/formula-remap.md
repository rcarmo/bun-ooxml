# Remap static references after insertion

`insertFormulaReferences(source, contextSheet, change)` returns a formula string
with supported A1 references shifted for one row or column insertion. It does not
insert worksheet cells or change any workbook package.

```ts
import { insertFormulaReferences } from 'bun-ooxml';

const formula = insertFormulaReferences(
  'IF(A1="A2",A2,Other!A2)',
  'Main',
  { axis: 'row', at: 2, count: 1, sheet: 'Main' },
);
// IF(A1="A2",A3,Other!A2)
```

`at` is a one-based insertion coordinate. References on or after it shift by
`count`. Unqualified references use `contextSheet`; qualified ones use their
decoded sheet name. Matching uses JavaScript `toLowerCase()` on those names,
without locale-specific comparison, trimming or Unicode normalisation. This is
a bounded API policy, not full Excel sheet-name identity validation.

Both range endpoints shift independently, retaining their order and absolute
flags. `$` does not prevent a structural insertion shift. If either endpoint
changes, both endpoints are rendered with uppercase column letters; their row
numbers and dollar markers remain structurally equivalent. The original sheet
qualifier is copied exactly, including quotes and doubled apostrophes. A single
cell stays a single cell, and an explicit repeated-endpoint range stays a range.

Unchanged references, literals, operators and surrounding spaces retain their
original bytes. A no-op returns the original string after full analysis. The
[static analyzer's grammar](formula-analysis.md) applies even when the insertion
would affect no references, so dynamic calls, names and malformed syntax still
refuse.

## Arguments and limits

`change` must be plain data containing only `axis`, `at`, `count` and `sheet`.
Accessors, inherited fields and unknown properties refuse. `axis` is `row` or
`column`; `at` and `count` are positive integers within that axis's grid limit:
1048576 rows or 16384 columns. The sum of `at` and `count` is not a worksheet-size
check. Only referenced coordinates that actually shift are tested for overflow.

Context and target sheets must be nonblank, well-formed strings of at most 255
UTF-16 code units, without control characters or external/3D-reference punctuation.
They are decoded names, not quoted formula tokens. This limit is an editor policy;
no workbook membership or Excel naming-rule check is performed.

Input retains the analyzer's limits. Output is limited to 1 MiB of UTF-8. Invalid
arguments and coordinate overflow throw `xlsx-formula-remap-unsupported`;
unsupported formula syntax uses `xlsx-formula-unsupported`, and resource limits
use `xlsx-formula-limit`. A late refusal returns no partial expression. Input
strings and insertion options are unchanged. These checks do not cover saved-file
rollback or peak process memory.

## Checks and boundaries

The shared cases assert five exact expressions and seven refusals. One further
shared case runs a finite 288-expression matrix, checks reparsable reference
slices, byte-identical unaffected remaps and `SUM(...)` span shifts. It is one
case with 288 iterations, not a fuzz campaign. Native controls also cover Unicode
qualifiers, absolute/reversed references, late overflow, output growth and
corrupted matrix results.

The [shared static-formula profile](../../references/fixtures-ooxml/contracts/go-formula-references.md)
defines these examples. No formula evaluation, workbook structural mutation,
cache update, calculation-engine compatibility or complete ECMA formula-language
conformance is established by string remapping.
