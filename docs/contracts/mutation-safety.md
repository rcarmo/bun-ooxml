# Shared mutation workflows

The reference checkout owns `workflows/mutation-safety.feature` and
`contracts/mutation-safety.json`. Eight scenarios expand to 19 cases over four
fixture policies. Each policy resolves a canonical `assetId`; reusable bytes are
stored only in `fixtures/<format>/<scenario-group>/`.

`features/shared.json` selects the central feature. The Bun runner applies
lifecycle tags in memory and compiles it with the official Gherkin compiler.
Bindings call `patchOffice` and assert saved output, receipt counts and custody.
No feature copy or generated expanded-case catalogue is stored locally.

The contract records readback facts, one member-hash map and an allow-list of
changed members. Package membership must remain exact; all members outside the
allow-list retain their hashes. `loadMutationFixtures()` resolves file identity
from the root manifest and derives the preserved set without duplicating metadata.
Original source lineage and transformations belong to the root asset's
`derived-from` origins, separate from the derived file's hash.

`scripts/shared-contracts.ts` verifies workflow/contract hashes, fixture readback,
scenario IDs, expanded case keys, typed mutation inputs and preservation policies.
Its input checks do not execute mutations. `make check` runs native bindings
separately and writes current results to `artifacts/acceptance.json`.

The historical revision `ooxml-shared-contracts-v2` and outcome interchange version
2 are identifiers, not directory layouts. The root manifest seals workflow and
contract files; the pinned Git commit seals the whole reference checkout.

Planned/unmapped contracts receive no execution credit. Native-library results,
server-call results and transport checks remain separate. Full cross-language
outcome export and complete behaviour-catalogue reconciliation are unfinished.
