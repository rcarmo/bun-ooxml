# Shared fixture references

The shared specification in `references/fixtures-ooxml` defines expected behaviour
and supplies reusable validation fixtures. The manifest is schema 2; storage is
grouped by format and scenario purpose, with one file per unique SHA-256.

## IDs and physical files

A fixture entry contains `id`, `path`, `bytes`, `sha256`, `role`, `format`,
`scenarioGroup`, `origins`, `aliases` and `scenarioIds`. IDs have the form
`fixture-<full-SHA-256>`. Paths are relative to the shared repository, for example
`fixtures/xlsx/mutation-safety/cross-sheet-cache.xlsx`. Short hash suffixes only
disambiguate distinct same-named fixtures; consumers never construct these paths.

The same file may serve several scenarios. Its ID is reused; neither another
origin nor another scenario creates a physical copy. `aliases` retains historical
paths for provenance. Test and development lookup uses IDs, not those aliases. Required notices are deduplicated under the shared `notices/`.

## Bun lookup

`scripts/fixture-inputs.ts` resolves IDs through the manifest and checks schema,
metadata, safe relative paths, file size and SHA-256. Symbolic fixture targets or
escaped paths refuse. The helper is test/development tooling; it is not a runtime
Office editing dependency.

`scripts/references.ts` verifies the pinned commit, annotated release tag, root
manifest seal, all asset hashes and a clean Git checkout. Every tracked file is
also compared with its raw Git blob, regular-file type, executable mode and
non-symlink path. This catches facts and workflows hidden by `assume-unchanged`,
`skip-worktree` or disabled mode tracking; validation does not modify index hints.
Missing references fail with submodule initialisation guidance. A changed fact or
workflow also fails even if every fixture hash still matches.

`contracts/mutation-safety.json` records allowed member changes for its fixtures.
`workflows/mutation-safety.feature` defines the corresponding scenarios.
`features/shared.json` selects the shared workflows used by Bun acceptance tests.
File identity and derivation are resolved through the root manifest.

The package-admission and semantic-diff profiles under `workflows/package/`, and
`workflows/xml/comparison.feature`, have no Bun bindings. Their presence in the
specification does not mean the Bun APIs implement them.

## Alternate reference checkout

Checks normally use the gitlink and `references/fixtures-ooxml.pin.json`.
To validate a different clean checkout, supply both overrides:

```sh
OOXML_FIXTURES_ROOT=/absolute/path/to/clean/candidate \
OOXML_REFERENCE_PIN=/absolute/path/to/candidate.pin.json make check
```

The override pin requires `mode: "candidate"`, the repository URL, exact commit
and root-manifest seal, with no release tag. Never write test output into either
reference checkout.
