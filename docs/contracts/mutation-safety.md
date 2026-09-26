# Shared mutation workflows

The shared specification defines mutation safety in `workflows/mutation-safety.feature`
and `contracts/mutation-safety.json`. Eight scenarios expand to 19 cases over four
fixture policies. Each policy resolves an `assetId`; reusable bytes are
stored only in `fixtures/<format>/<scenario-group>/`.

`features/shared.json` selects the central feature. The Bun runner applies
lifecycle tags in memory and compiles it with the official Gherkin compiler.
Bindings call `patchOffice` and assert saved output, receipt counts and preserved
package parts.
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

The root manifest verifies workflow and contract files; the Git revision identifies
the complete reference checkout. Acceptance results cover the selected scenarios
and do not establish equivalent behaviour in other Office libraries.
