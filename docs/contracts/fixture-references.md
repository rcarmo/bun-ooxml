# Shared fixture references

All reusable fixture bytes and common behavioural references come from the
`references/fixtures-ooxml` submodule. The manifest is schema 2; storage is grouped
by format and scenario purpose, with one file per unique SHA-256.

## IDs and physical files

A fixture entry contains `id`, `path`, `bytes`, `sha256`, `role`, `format`,
`scenarioGroup`, `origins`, `aliases` and `scenarioIds`. IDs have the form
`fixture-<full-SHA-256>`. Paths are relative to the shared repository, for example
`fixtures/xlsx/mutation-safety/cross-sheet-cache.xlsx`. Short hash suffixes only
disambiguate distinct same-named fixtures; consumers never construct these paths.

The same file may serve several scenarios. Its ID is reused; neither another
origin nor another scenario creates a physical copy. `aliases` retains historical
paths for provenance and one-time migrations. Runtime/test lookup uses IDs, not
those aliases. Required notices are deduplicated under the shared `notices/`.

The grouped baseline contains 115 unique fixture files in 32 groups, totalling
6,451,099 bytes. The complete manifest has 125 asset entries, including notices
and shared workflow metadata. These counts have different meanings from the 74
logical package inputs exercised by the corpus regression tests.

## Bun lookup

`scripts/fixture-inputs.ts` resolves IDs through the manifest and checks schema,
metadata, safe relative paths, file size and SHA-256. Symbolic fixture targets or
escaped paths refuse. The helper is test/development tooling; it is not a runtime
Office editing dependency.

`scripts/references.ts` verifies the pinned commit, annotated release tag, root
manifest seal, all asset hashes and a clean Git checkout.
Missing references fail with submodule initialisation guidance. A changed fact or
workflow also fails even if every fixture hash still matches.

`contracts/mutation-safety.json` references four fixture IDs and records their
member-preservation policies. `workflows/mutation-safety.feature` defines the 19
expanded cases. File identity and original derivation are resolved through the
root manifest; no separate fixture manifest, pack or generated cases are stored.

XML parsing and QName acceptance read `workflows/xml/parsing.feature` and
`workflows/xml/names.feature` directly. The old local XML features are removed;
scenario IDs and all 11 expanded cases are unchanged. Python comparison has a
separate `workflows/xml/comparison.feature` profile, without a Bun binding.
Reviewed native-test snapshots in `ledgers/consumers/` record overlaps and gaps;
only the Bun acceptance report records executed Bun cases.

## Candidate checks

Release checks use the gitlink and `references/fixtures-ooxml.pin.json`.
Pre-release testing requires both explicit overrides:

```sh
OOXML_FIXTURES_ROOT=/absolute/path/to/clean/candidate \
OOXML_REFERENCE_PIN=/absolute/path/to/candidate.pin.json make check
```

A candidate pin has `mode: "candidate"`, repository URL, exact commit and the root
manifest seal. It cannot claim a release tag. Candidate results do not establish
that a released recursive clone works; run the default checks again after tagging
and repinning. Never write test output into either checkout.
