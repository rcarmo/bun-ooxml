# Native behaviour catalogue

`native-test-inventory.json` records the current Bun test declarations from
`tests/unit/**/*.test.ts`. Schema 2 contains 389 declarations in 42 files. This is an
AST denominator, not a runtime-expanded case count or proof of semantic coverage.
It records whole-file hashes, declaration locations, suite/title identities,
inline bodies, direct matcher expressions, call names and review reasons.

82 declarations have dynamic titles, parameterisation, loops, conditions or
lifecycle modifiers requiring review. 113 declarations need some assertion or
expansion review, including tests whose assertions are in helpers. These groups
overlap. The inventory does not evaluate helper bodies or expand runtime tables.
All source files in the discovery glob are hashed, including files with no tests.
The snapshot also hashes 29 other TypeScript files under `tests/`, so test-helper
and acceptance-binding changes invalidate it. Dependencies outside `tests/` are not
hashed by this inventory; runtime behaviour is covered by separate test runs.
Three declarations contain deferred matcher syntax in nested functions; that
syntax is listed separately and does not count as a direct assertion.

## Drift gates

```sh
bun scripts/test-inventory.ts
bun scripts/mapping-reconciliation.ts
make check
```

`make check` and `bun run check` compare committed outputs with regenerated data.
Adding/removing a declaration or changing a discovered file fails until the
inventory is refreshed. The check does not rewrite reports. Duplicate test
identities and unresolved registrations refuse. Named and namespace `bun:test`
imports are recognised; indirect aliases, tagged tables, non-inline test/suite
callbacks, deferred helper registrations and detected binding shadowing need
explicit parser/review work. Transparent TypeScript wrappers are unwrapped before
alias checks. Destructured, default/rest and catch bindings are inspected for
shadowed imports.

This is syntax analysis, not a general TypeScript execution model. Discovery is
limited to the unit-test glob and supported registration syntax. Integration and
oracle workflows have separate reports and remain outside this denominator.
No inventory row grants execution credit.

## Historical candidate mappings

The five `staging-*.json` candidate groups retain 227 earlier declaration IDs.
`staging-reconciliation.json` finds all 227 identities in the current source and
lists 162 newer declarations without historical staging. The report itself is
excluded from candidate discovery. These counts describe only those five files,
not the separate canonical consumer mappings.

The old candidates have no source pins. A matching test ID does not prove that
its body, helper assertions or fixture semantics are unchanged. Every surviving
row stays `identity-present-needs-review`, with source equivalence unverified and
execution credit false. Missing IDs would remain explicit rather than disappearing.
Each source candidate must be read and compared before central reconciliation.

Shared `ledgers/consumers/bun-xml.json`, `bun-package.json` and `bun-comments.json`
contain their own source-pinned, bounded mappings. They supersede corresponding
historical prose only for the assertions they explicitly cover. Comment mappings
remain partial; archive custody, direct flags and other native assertions can
exceed the shared scenario wording.

## Source-pinned outcome mappings

[`slide-order-mappings.json`](slide-order-mappings.json) maps all 13 slide-order
native declarations to canonical scenario IDs, literal native assertions,
reviewed outcomes and explicit gaps. Its wrapper lists the 21 canonical case
keys; other declarations retain scenario-level partial mappings because their
anchor, encoding, constructor, rollback and invalid-input variants lack exact
canonical rows.

[`outcome-reconciliation.json`](outcome-reconciliation.json) retains the full
389-declaration denominator: 13 have bounded mappings and 376 are unmapped by
this ledger. All mappings remain partial and carry `executionCredit: false`.
Each row records `scenario-only` or `explicit-case-keys` link granularity.
Body-loop review flags remain present, and the runtime leaf count is unknown.
The 296 acceptance cases still supply the separately recorded execution result.

`bun scripts/outcome-mappings.ts --check` fails on stale reviewed source hashes,
missing scoped declarations, duplicate or unknown IDs, missing or extra direct
literal assertions,
and missing outcome/gap descriptions. The reviewed source set includes the
native test, acceptance helper, PPTX implementation, canonical feature, Gherkin
runner and inventory parser. It is not a transitive dependency closure. The gate
checks integrity; it cannot determine whether prose accurately describes an
assertion. Review source changes before updating pins. `make check` runs this
gate without awarding additional execution credit.

These consumer mappings add no behaviour IDs or duplicate Gherkin. The central
reference remains the source of canonical scenarios. Wider Bun/Go/Python mapping
and runtime-leaf reconciliation is still open.

## Reconciliation work

Review concrete Then predicates, helpers, negative paths and parameter rows.
Keep weak substring/count assertions separate from exact equality or preservation.
Record the current file hashes and gaps when proposing shared IDs. Preserve
operation differences between consumers. Shared Gherkin and facts live centrally;
local mappings, inventories and native results remain consumer-specific.

Current acceptance executes 296 implemented cases selected from the shared
reference. Neither 389 declarations nor 227 candidate IDs increases that result or
closes the full-format backlog.
