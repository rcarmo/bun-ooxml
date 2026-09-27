# Tests and specification coverage

The shared ECMA specification defines expected OOXML behavior. Tests check the
implemented operations against those requirements. Application output is useful
for compatibility checks, but does not override the standard.

## Find results

| File | Contents |
|---|---|
| `artifacts/acceptance.json` | Results for every executed Gherkin case and step |
| `native-test-inventory.json` | Unit-test declarations, assertions and source hashes |
| `outcome-reconciliation.json` | Connections between tests and shared scenarios, including missing assertions |
| `*-check.json` | Recorded validation commands, source versions and results |

The current Gherkin inventory is in
[behaviour coverage](../contracts/behaviour-coverage.md). Planned scenarios are not
run. A unit-test declaration may contain loops or call helpers, so its presence
alone does not count as a tested requirement.

## Maintain the reports

```sh
bun scripts/test-inventory.ts
bun scripts/mapping-reconciliation.ts
bun scripts/outcome-mappings.ts
make check
```

Review test and requirement changes before refreshing reports. `make check`
compares the saved reports with current source and rejects missing declarations,
changed source hashes, duplicate identities and unsupported registrations.

The inventory reads `tests/unit/**/*.test.ts` and hashes supporting TypeScript
files under `tests/`. It records direct assertions separately from assertions in
nested functions. It does not execute parameter tables or analyse arbitrary
TypeScript callbacks. Integration and Office-application checks have separate
reports.

The ZIP32 acceptance tests execute checksum, typed-refusal and configured-limit
examples against the native ZIP API. They share the original unit-test archive
builder; separate tests inspect mutated header fields and reject incorrect error
codes, messages, fabricated success and changed input buffers. Limit refusals
check API results, without measuring allocation timing or peak memory.

The OPC custody cases use the shared three-member package example, without a
Word schema or renderer. They check detached buffers, UTF-16LE member bytes,
transaction callback behaviour and existing-file/symlink refusal. Edited archives
are reopened and unrelated member payloads compared; these comparisons do not
require unchanged ZIP headers after an edit. Save tests read destination bytes
from disk, and the runner removes temporary directories even after a failed case.
Concurrent path replacement is outside these tests.

Selected DOCX document-model cases check an empty body, eight authored table
sizes and a nine-cell table saved to disk and reopened. The size cases inspect
in-memory getters; only the saved-table case checks disk readback. Controls alter
each saved cell position independently and verify temporary-directory cleanup.
Bun throws on out-of-range cell access, so the shared nil-return policy stays
planned, along with merge-property operations. The row-count case uses bounded
empty-row append/insertion/deletion, checks 3/4/3 in memory and verifies an invalid
deletion error. Separate native tests check row text/order and package custody
after save/reopen. The first-row text case separately compares every cell in a
2x2 table and requires the first-row result to contain exactly two ordered values;
controls reject wrong values, reversed order and missing or extra array members.

## Requirement mappings

The slide-order, effective-formatting, XML-value, PPTX-core, DOCX-model and formula
reports associate literal assertions with shared scenarios. They also list gaps, such as an untested result field or
an input variant missing from the scenario. Each report has its own source hashes
and required files; one cannot supply a missing reference for another.

The XML ledger covers `tests/unit/xml.test.ts`, including parser, escaping and
edit predicates. Its metadata-map link uses a different fixture from the shared
scenario, so it stays scenario-only. Positive and negative edit declarations
also map separately. `tests/acceptance/xml-values.ts` executes the shared value
inputs and checks each result; negative controls corrupt values and metadata to
check that these assertions fail. The ledger itself does not execute tests.

The [PPTX-core ledger](pptx-core-mappings.json) maps three suites to text, notes
and no-op-save scenarios. The six original `pptx.test.ts` records are unchanged.
Cross-run editing compares two unrelated saved payloads; the no-op acceptance
case compares file bytes and its reopened archive with the original fixture.

The two notes suites add 27 declarations and 109 direct assertion expressions.
Two positive wrappers link to exact case keys; the other 25 mappings link at
scenario level. Wrappers check aggregate results and counts, with semantic checks
in the pinned bindings. The direct splice test reopens bytes; the shared splice
case saves to a path. The string-refusal test never supplies raw `FF`, while the
compound wrapper does. Its identical-text no-op uses the string entry point.
Native UTF-8 tests separately check byte no-ops and saved Unicode values.

Other differences include a bold first-run template where the shared case uses
explicit false, and late metadata changes that stale an existing target before
new topology can be validated. These gaps remain in the individual mappings.
The mappings record no rendering result or new execution credit.

The [DOCX-model ledger](docx-model-mappings.json) records all 130 declarations and
683 direct assertion expressions in ten suites: append-run, cell properties,
row headers, paragraph text, body insertion, table-row edits, row-text reads,
core properties, document properties and table styles. They link to twelve
existing document-model scenarios. Two direct run tests and ten acceptance-wrapper
tests link to exact case keys. The other 118 mappings
link at scenario level because they use different inputs or add refusal,
preservation, encoding and resource-limit checks. Every mapping retains partial
status and lists its gaps.

The wrapper tests check aggregate results and failure counts; their semantic
predicates live in the pinned step bindings. Saved-formatting corruption changes
the reopened in-memory document, not the file on disk. Most native row-header
examples use 2x2 tables, while the shared case uses 3x2. Cell-property tests combine
values on 2x2 tables; the shared width/border case uses 1x1. Row-edit tests use a
2x2 input while the shared count sequence uses 2x3. Their disk check reopens only
the final state, without verifying each intermediate count after save. Row-text
roundtrips compare array results; they do not assert all four direct cell getters
after reopen. The paragraph-text loop reuses one paragraph across values, while
the shared outline creates independent targets. These differences stay explicit.
Neither exact-case links nor complete suite enumeration add execution credit. Dynamic loops retain their review flags, and the runtime leaf count is
unknown. The source pins cover the reviewed files, not their transitive imports.

The core-property scenario supplies fifteen fields but compares only title,
creator and subject from a cached getter result. Its native path-roundtrip test
compares all fifteen; relationship and content-type checks occur before save.
The document-property scenario checks direct title-page and background values
in memory. Its native tests add path readback, encoding, custody and refusal
checks. The table-style scenario compares captured strings before and after
assignment, converting an absent ID to an empty string. The native API returns
`undefined` for absence and writes a reference without resolving a style or
predicting its appearance. None of these three shared getter scenarios saves or
reopens a file. Their wrapper tests assert aggregate results and failure counts;
separate native tests provide the saved-file checks.

The [formula ledger](formula-references-mappings.json) records 31 declarations
and 127 direct assertion expressions from the range, analysis and remapping
suites. Nine mappings link to exact case keys, including three acceptance
wrappers. The other 22 link at scenario level. The analysis-only 288-expression
matrix lacks the shared remapping predicate; the remapper's matrix checks all
three outcomes. Both are finite loops, with no fuzzing or runtime-leaf count
inferred. Literal-punctuation and remap-refusal predicates split across native
tests retain that distinction. Range-wrapper planned counts refer to its isolated
selection, not the full acceptance run. No mapping establishes worksheet editing,
formula calculation or saved-workbook behaviour.

Unresolved registrations, missing assertions or conflicting source records fail
validation. The descriptions still need human review: matching an assertion's
source text cannot establish that it tests the intended specification requirement.

Older `staging-*.json` files are retained as historical notes. Their matching test
names do not establish current behavior. Full format coverage and test-to-clause
mapping remain incomplete.
