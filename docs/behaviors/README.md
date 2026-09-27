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

## Complete shared inventory

The project loads all 46 canonical features from the pinned workflow registry.
Only explicitly selected scenario IDs are implemented; all others stay planned.
Acceptance JSON includes `coverage.shared`, `coverage.localOnly` and the catalogue
path/hash. The coverage table reports the same independent denominators:

- Shared: 229 scenarios / 562 cases; 525 implemented and 37 planned.
- Local-only: nine scenarios / eleven cases, all planned.
- Combined: 238 scenarios / 573 cases; 525 implemented and 48 planned.

This closes an inventory omission of thirteen template cases. Earlier release
notes below describe their historical loaded counts. The executed case-key set
and all 472 mapping records are unchanged; visibility does not award a pass.

## Format and operation layout

Shared v0.22 groups features under DOCX, PPTX, XLSX, package and XML operations.
The [shared operation index](../../references/fixtures-ooxml/workflows/README.md)
replaces the former `native/`, mixed document-model and root workflow files.
The shared migration ledger preserves all scenario IDs, compiled cases and tags.

Bun keeps the same 525 implemented shared case identities. Test wrappers select
IDs across current files instead of activating every scenario in a file. Source
pins were renewed after reviewing the moves. Of 472 mapping records, 470 are
unchanged; two wrapper assertions now account for two PPTX features and zero
unselected cases in the narrowly scoped run-authoring wrapper. No document
runtime code changed. The shared `CATALOGUE.md` lists remaining runtime-specific
API contracts and wording; file grouping alone does not resolve those differences.

## Portable action wording

Shared v0.23 uses runtime-neutral actions in presentation notes, presentation
archive preservation, relationship namespaces and ZIP64. Bun bindings accept the
new wording only. A bounded execution check runs all 16 scenarios / 18 cases;
default acceptance remains the same 525 cases with 35 planned. All 472 mapping
records and document-runtime sources are unchanged; affected source pins were
renewed for the feature and binding edits.

The shared wording ledger records exact substitutions. Historical IDs retain
their origin names. Existing-notes and archive-noop profiles describe operations;
ZIP64's exact error codes and safe-integer bound remain API compatibility policies.
Other runtime-specific contracts listed in the shared catalogue still need review.

Shared v0.24 also neutralises package custody and ZIP32 actor names across 18
scenarios / 38 cases. Exact `OoxmlError`, code/message strings and option keys
remain API compatibility predicates; synchronous callback/thenable behaviour is
explicitly JavaScript-specific. No assertion or runtime code changed. New bindings
reject the retired actor wording, and existing corrupt-buffer, false-success,
callback-flag, thenable-identity and destination checks still run. All 472 mapping
records and the 525 default case identities are unchanged.

Shared v0.25 neutralises Word setup wording and replaces source-runtime profile
labels for 34 IDs / 75 cases. Bun still selects 28 of those IDs / 65 cases; the
remaining six profiles / ten cases stay planned. Heading classification, nil
cell access, simultaneous in-memory effects, merge getters, tracking toggles and
tool hints are not activated by a label change. Getter names, values, saved-output
assertions, the 525 default identities and all 472 mappings remain unchanged.
The shared Word wording ledger and its exact before/after guards record the edits.

Shared v0.26 neutralises actors in XML snapshot editing and static formula
references across 19 IDs / 60 cases. Exact XML bytes, JSON/table inputs, namespace
combinations, UTF-8 reference spans, grammar refusals and string remap outputs are
retained. The matrices still count as one canonical case each. Bindings accept
only the new wording; corruption controls and unchanged-default-identity checks
remain active. This adds no workbook calculation, schema or rendering coverage,
and changes none of the 472 mapping records.

Shared v0.27 completes the catalogue-wide actor/profile wording pass: comment
and template APIs use operation labels, XML escaping/error profiles lose their
origin prefixes, and XLSX creation keeps fixture provenance without a Bun actor.
The changed 24 IDs / 25 cases do not add Bun bindings. Nine authored-comment cases
stayed planned; thirteen template cases were outside Bun's loaded inventory at
that release and are now included as planned by the complete catalogue loader.
Only the two already implemented XML profile IDs remain active. All 525 default
case identities and 472 mapping records are unchanged.

The shared catalogue's weak-outcome review records dictionary-only responses,
non-vacuous-filter gaps and missing cache/authoring custody checks. Neutral words
do not turn those API observations into complete semantic analysis or parity.

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

The mapping reports associate literal assertions with shared scenarios. They
also list gaps, such as an untested result field or
an input variant missing from the scenario. Each report has its own source hashes
and required files; one cannot supply a missing reference for another.

The XML ledger covers `tests/unit/xml.test.ts`, including parser, escaping and
edit predicates. Its metadata-map link uses a different fixture from the shared
scenario, so it stays scenario-only. Positive and negative edit declarations
also map separately. `tests/acceptance/xml-values.ts` executes the shared value
inputs and checks each result; negative controls corrupt values and metadata to
check that these assertions fail. The ledger itself does not execute tests.

The [package-admission ledger](package-admission-mappings.json) maps 11
declarations and 36 direct expressions across the shared ZIP-admission and
XML-member-admission features. Both feature files are required source pins.
One aggregate wrapper links to all 13 cases; the other ten records remain
scenario-only, including positive controls that extend refusal contracts.
All 461 prior records remain unchanged.

Admission validates ZIP structure and XML syntax without requiring an OPC graph.
The BZIP2 sample contains fixed, offline-compressed bytes, but admission rejects
its method without decoding them. Depth and exact-budget tests assert observed
success/refusal, not timing, memory use or an allocation trace. Encoding custody,
opaque members and buffer subviews are bounded native extensions; no mapping
adds filesystem, Office-schema or rendering evidence.

The [ZIP32 ledger](zip32-mappings.json) maps 24 declarations and 106 direct
assertion expressions across reader/writer tests and binding controls. One wrapper
links to 20 shared cases; its four planned siblings describe isolated selection,
not the full acceptance run. The other 23 records are scenario-only. All 422
previous mappings remain unchanged.

Existing refusal and aggregate-failure predicates now sit in their test bodies;
the two suites still execute 186 assertions. Literal-expression inventory
collapses repeated expressions, while nested fixture-signature and source-match
checks remain separate. The 74-archive corpus loop is admission smoke, not exact
member readback. Method 12 inputs contain plain bytes rather than BZIP2 streams.
Bounds tests observe errors, not allocation timing. Native refusal tests do not
check input custody; the pinned acceptance bindings do. No mapping implies
filesystem persistence or complete OPC validity.

The [ZIP64 ledger](zip64-mappings.json) maps 15 declarations and 46 direct
assertion expressions, all at scenario level. Eleven refusal tests now carry
their existing error-class/code/message assertions directly; moving them out of
the helper preserves the suite's 50 runtime assertions. The two writer tests
still use a layout helper with two signature assertions per call, separate from
their direct expressions. All 407 earlier mapping records are unchanged.

Most inputs are tiny archives with forced ZIP64 metadata. The 65,535-entry test
uses empty payloads and checks the count plus first/last values; it is not a
multi-gigabyte data or offset test. Signed 64-bit descriptors are exercised, but
other builder branches do not imply coverage. There is no native wrapper or
caller-budget refusal in this suite, and the forced writer test does not rewrite
a changed member as the shared roundtrip requires. Structural refusals are
associated with directory preflight scenarios as extensions, not exact cases.
No mapping claims allocation timing, file persistence or full ZIP64 parity.

The [OPC custody ledger](opc-custody-mappings.json) maps 15 declarations and
66 direct expressions. One wrapper links to 11 selected cases; the other 14
records remain scenario-only. Its three planned siblings describe isolated
selection, not global coverage. Equivalent helper inlining preserves the two
suites' 231 runtime assertions, and all 446 prior mapping records are unchanged.

Native corpus coverage compares each of 74 archive serializations exactly, but
does not count the two source corpora separately or reopen each result a second
time. The UTF-16 test uses the ZIP layer; the binding adds OPC reopen and unrelated
member checks. Async-function refusal happens before the body runs; a synchronous
callback can return an untouched thenable. This is not general rollback evidence.
Physical save-refusal tests check existing files and symlink targets; the binding
also checks link identity, and dedicated controls verify temporary-file cleanup.
None establishes trusted-ancestor safety, concurrent path replacement or Word
schema validity for the deliberately minimal three-member package.

The [OPC graph ledger](opc-graph-mappings.json) maps 15 declarations and 56
direct expressions from graph-edit and byte/content-type diff tests. All links
are scenario-only: neither suite contains a shared acceptance wrapper. All 392
previous records remain unchanged. Native graph tests reopen serialized bytes,
not paths, and individual rows state which payloads or metadata are compared.
The first add test checks the relationship target, MIME and main document bytes,
but not the new opaque payload itself.

Diff tests compare hashes, lengths, MIME and exact category arrays. Some build
changes with low-level `set`/`delete` and manual content-type edits; those tests
do not prove guarded graph-edit behaviour. The metadata-only test changes a ZIP
comment, not every possible metadata field. Orphan traversal, name allocation
and relationship identity are bounded extensions of the shared graph scenarios.
None of these mappings adds semantic XML or Office-rendering credit.

The [XML-comparison ledger](xml-comparison-mappings.json) maps 14 declarations
and 42 direct expressions. The [package-comparison ledger](package-comparison-mappings.json)
maps 12 declarations and 32 expressions. Their two aggregate wrappers link to
ten XML cases and one package case; the other 24 records remain scenario-only.
All 347 prior records are unchanged.

XML equivalence is conservative: namespace-sensitive values, significant text,
comments and processing instructions remain meaningful. Unsafe identical inputs
can compare false. This is not canonical XML, schema validation or signature
verification. Package comparison first admits both archives, then classifies
member payloads; it does not require an OPC graph. Exact payload equality stays
separate from XML equivalence and from ZIP-byte equality. The existing OPC byte
diff still reports prefix-only rewrites as changed.

The package binding checks four lists, not `unchanged`. Native controls check
additional classifications and input custody. The category-union test has only
one member per populated category, so it does not independently test within-list
sorting. The XML depth test provides accepted/refused boundary examples, not
timing or allocation evidence. No mapping establishes filesystem readback or
Office rendering.

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

The [PPTX text-box ledger](pptx-text-boxes-mappings.json) maps 11 declarations
and 40 direct assertion expressions. The [DOCX page-layout ledger](docx-page-layout-mappings.json)
maps 13 declarations and 44 expressions. Each has one aggregate acceptance
wrapper linked to all cases in its shared feature (23 and 22 respectively); the
other 22 records are scenario-only. All 295 earlier mapping records are unchanged.

Their positive bindings reopen serialized bytes, not paths. Text-box geometry
checks compare integer EMU attributes; page-layout checks compare direct twip
values. Neither establishes rendered geometry or pagination. Separate disk tests
check text-box paragraph/flag/extension ordering and final page geometry with an
earlier section XML substring. These disk tests do not assert all package bytes.
The page-layout UTF-16 test checks BOM/declaration after serialization, but its
paragraph getter is still live and it has no reopened geometry assertion. The
real-presentation text-box test also reopens bytes, despite its saved-file wording.
All these differences remain explicit in the ledgers; independent Office checks
are separate evidence, not credited by these source mappings.

The [paragraph-style ledger](docx-paragraph-styles-mappings.json) maps 17
declarations and 62 direct expressions; the [style-authoring ledger](docx-style-authoring-mappings.json)
maps 11 declarations and 43 expressions. One aggregate wrapper per ledger links
to its 19 or 24 shared cases. The other 26 records remain scenario-only, and all
319 prior mapping records are unchanged.

Both bindings reopen serialized bytes rather than paths. Assignment custody
checks selected main-part fragments and complete unrelated member payloads, not
an exact main-XML edit diff. Authoring checks definition metadata and old child
substrings before a separate selection step changes the paragraph reference.
Two native disk tests provide bounded path readback. Several encoding tests use
live getters or omit new-definition assertions; those limits remain explicit.
Direct references and basedOn links do not establish inherited formatting, and
these ledgers add no execution or rendering credit.

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

The [DOCX-anchor ledger](docx-anchors-mappings.json) records 28 declarations and
122 direct assertion expressions from the body-anchor and body-map suites.
Two acceptance wrappers link to exact case keys; the other 26 records link at
scenario level. Their 251 predecessor mappings remain unchanged. Wrappers check
aggregate results; semantic predicates live in the pinned bindings. Native
listing and encoding tests reopen bytes, while the direct map test saves and
reopens a real path. Insertion formatting is checked before save, with paragraph
order and unrelated payloads checked after reopen. Negative map-custody controls
require an assertion failure at the request step, not a serialization error.

Headings use explicit outline levels, not inherited style names. Map sections
count heading markers rather than Word section breaks; placeholders are bounded
literal occurrences in direct body paragraphs, excluding table cells. No mapping
links to the planned tool-hint case. These records add no execution, transport or
rendering credit; dynamic and nested assertions retain their inventory flags.

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

The [XLSX-styles ledger](xlsx-styles-mappings.json) covers 16 declarations and
52 direct assertion expressions. One wrapper links to all 27 selection/refusal
case keys; its two direct assertions check aggregate results, while the pinned
bindings check semantic outcomes. Positive bindings reopen bytes, not a saved
path. The other 15 records remain scenario-only, including a separate disk-save
test that compares the formula cell and styles payload after reopen.

The UTF-16 test checks encoding markers and cell text, but does not assert the
reopened style index. Protected absent-index and existing-index tests are mapped
separately: the former's zero/one requests are assignments, not same-index no-ops.
Row and column default markup is preserved without computing its effective
formatting. These mappings do not cover the independent SDK style reader or new
style creation. All 279 prior mapping records remain unchanged.

The [DOCX-comments ledger](docx-comments-mappings.json) maps 19 declarations
and 60 direct expressions. One aggregate wrapper links to the 14 existing-comment
cases; the other 18 records are scenario-only. All 373 previous records remain
unchanged. None links to the nine Python comment-authoring, filtering or
reply-to-root workflows, which have different policies.

The binding opens the pinned file, then reopens edited bytes in memory. It checks
that only the existing extension part changes, the reply stays open with its
parent link intact, and reopening the target restores member payloads. This does
not assert restored ZIP bytes. A separate native disk test checks the saved done
flag and comments body part, not every package member. Refusal bindings require
an `OoxmlError` instance and exact archive custody, without a specific error code.
Native codec coverage is UTF-16LE only; unsupported bodies remain inspectable as
findings, while resolution refuses. No mapping implies comment creation or
rendered-comment equivalence.

The [XLSX-comments ledger](xlsx-comments-mappings.json) maps 19 declarations and
72 direct assertion expressions from the native inspector and binding controls.
Only the positive canonical wrapper has an exact-case link. Its five planned
cases refer to the isolated comment/VML feature, and its semantic checks live in
the pinned binding. Negative tests keep failure counts in their own bodies;
mutation callbacks and the binding-selection helper have separate nested/helper
assertions. Some malformed dependency controls fail while serialising their
baseline, before the inspector runs. Other controls reset that baseline to check
semantics independently of byte preservation. Native tests compare member arrays
when an invalid package cannot serialize. No mapping links to the five mutation
profiles or establishes VML shape validity.

Unresolved registrations, missing assertions or conflicting source records fail
validation. The descriptions still need human review: matching an assertion's
source text cannot establish that it tests the intended specification requirement.

Older `staging-*.json` files are retained as historical notes. Their matching test
names do not establish current behavior. Full format coverage and test-to-clause
mapping remain incomplete.
