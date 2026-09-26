# Shared fixture references

All reusable fixture bytes and common behavioural references come from the
`references/fixtures-ooxml` submodule. The manifest is schema 2; storage is grouped
by format and scenario purpose, with one file per unique SHA-256.

## IDs and physical files

A fixture entry contains `id`, `path`, `bytes`, `sha256`, `role`, `format`,
`scenarioGroup`, `origins`, `aliases` and `scenarioIds`. IDs have the form
`fixture-<full-SHA-256>`. Paths are relative to the shared repository, for example
`fixtures/docx/comments/comments-<hash-prefix>.docx`.

The same file may serve several scenarios. Its ID is reused; neither another
origin nor another scenario creates a physical copy. `aliases` retains historical
paths for provenance and one-time migrations. Runtime/test lookup uses IDs, not
those aliases. Required notices are deduplicated under the shared `notices/`.

The grouped baseline contains 115 unique fixture files in 32 groups, totalling
6,451,099 bytes. The complete manifest has 123 asset entries, including notices
and shared workflow metadata. These counts have different meanings from the 74
logical package inputs exercised by the corpus regression tests.

## Bun lookup

`scripts/fixture-inputs.ts` resolves IDs through the manifest and checks schema,
metadata, safe relative paths, file size and SHA-256. Symbolic fixture targets or
escaped paths refuse. The helper is test/development tooling; it is not a runtime
Office editing dependency.

`scripts/references.ts` verifies the pinned commit, annotated release tag, root
manifest seal, shared-pack seal, all asset hashes and a clean Git checkout.
Missing references fail with submodule initialisation guidance. A changed fact or
workflow also fails even if every fixture hash still matches.

The shared mutation fixture manifest uses `pathBase: "repository-root"` and
`assetId` references. Its four logical fixture IDs and 19 case identities retain
their existing meaning. No fixture bytes live inside the shared workflow pack.

## Candidate checks

Release checks use the gitlink and `references/fixtures-ooxml.pin.json`.
Pre-release testing requires both explicit overrides:

```sh
OOXML_FIXTURES_ROOT=/absolute/path/to/clean/candidate \
OOXML_REFERENCE_PIN=/absolute/path/to/candidate.pin.json make check
```

A candidate pin has `mode: "candidate"`, repository URL, exact commit and both
manifest seals. It cannot claim a release tag. Candidate results do not establish
that a released recursive clone works; run the default checks again after tagging
and repinning. Never write test output into either checkout.
