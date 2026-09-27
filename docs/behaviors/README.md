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

## Requirement mappings

The slide-order, effective-formatting, XML-value and PPTX-core reports associate
literal assertions with shared scenarios. They also list gaps, such as an untested result field or
an input variant missing from the scenario. Each report has its own source hashes
and required files; one cannot supply a missing reference for another.

The XML ledger covers `tests/unit/xml.test.ts`, including parser, escaping and
edit predicates. Its metadata-map link uses a different fixture from the shared
scenario, so it stays scenario-only. Positive and negative edit declarations
also map separately. `tests/acceptance/xml-values.ts` executes the shared value
inputs and checks each result; negative controls corrupt values and metadata to
check that these assertions fail. The ledger itself does not execute tests.

The PPTX-core ledger maps `tests/unit/pptx.test.ts` to the original four text/notes
scenarios and one path/byte-open/no-op-save scenario. Its aggregate acceptance
test checks report status and counts; the delegated step assertions are recorded
separately in source. Cross-run editing compares two unrelated saved payloads,
without a complete saved-member comparison. The no-op acceptance case saves to
a real temporary path and compares both its file bytes and reopened archive with
the original fixture. Rendering and newer notes-editing operations have separate
tests.

Unresolved registrations, missing assertions or conflicting source records fail
validation. The descriptions still need human review: matching an assertion's
source text cannot establish that it tests the intended specification requirement.

Older `staging-*.json` files are retained as historical notes. Their matching test
names do not establish current behavior. Full format coverage and test-to-clause
mapping remain incomplete.
