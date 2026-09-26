# Shared mutation workflows

The `references/fixtures-ooxml` submodule owns eight mutation scenarios, 19
expanded cases and the metadata for four unique input fixtures. Those fixture
bytes live in the shared `fixtures/<format>/<scenario-group>/` tree and are
resolved by `assetId`; the workflow pack contains no separate fixture copies.

`features/shared.json` selects the central feature. The Bun runner changes only
its lifecycle tags in memory before compiling it. Step bindings call native
`patchOffice` and assert saved outcomes, receipt counts and unrelated-part custody.

The grouped distribution uses schema 2 for its asset and fixture manifests.
The workflow revision remains `ooxml-shared-contracts-v2`; it is distinct from the
shared repository's release version and from the version-2 outcome interchange.
Stable case keys, fixture hashes and expected workflow results are unchanged.

`scripts/shared-pack.ts` checks the distribution seal, provenance, fixture IDs,
expanded cases and preservation allowances. Its inventory check does not execute
mutations. `make check` runs the native bindings separately and writes current
results to `artifacts/acceptance.json`.

Planned/unmapped contracts receive no execution credit. Native-library results,
server-call results and transport checks remain separate. The complete
cross-language outcome exporter and all-format behaviour catalogue are unfinished.
