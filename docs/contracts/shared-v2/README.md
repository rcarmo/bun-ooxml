# Deterministic shared Office contracts

The native shared workflow execution copy contains eight implemented scenarios
and 19 expanded cases; all pass through `patchOffice()`. The frozen export retains
its original planned tags. Four pinned inputs derive from the already imported Python MCP
revision `36ac406ad9d4bd3e7538b4bcc7aa2fb0e51cc943`. The pack is stored under
`pack/`; `manifest.json` pins its manifest, source revision and active IDs. The
v2 manifest hash is `4fb30e0d1a75e889985eceb0c6929dc59971089cc3bc692f18675f36dfeb81de`.

The six earlier audit IDs are archived under `../office-mutation/` with explicit
replacement mappings. The additional DOCX/XLSX previews and absent/existing output
variants expand coverage requirements. Historical IDs are not extra passes. Three
Bun additive obligations retain preview descriptions, exact match counts and an
independent style reader.

## What executes now

`scripts/shared-pack.ts` verifies every pack file, compiles the original feature
with the official Gherkin parser, and compares all expanded steps, arguments,
source locations and example values with the exported inventory. The active Bun
feature differs from the v2 pack only by replacing `@planned` with
`@implemented @bun`; the verifier refuses any other change. `stableCaseKey` uses
scenario ID plus sorted example values; compiler UUIDs and local file paths never
form cross-runtime identity. Every `value_json` table entry is decoded as JSON so
numeric 10 differs from string "10" and escaped newlines become actual newlines.

`scripts/shared-fixtures.ts` validates archive/member hashes, native readback
facts, relationship/content-type references, explicit and implicit style indices,
and the opaque sentinel's payload and root relationship. Preservation checks
reject added/deleted members and any payload change outside the allowed set.
Allowed parts still need semantic assertions; an allowance is only an upper bound.

Native bindings in `tests/acceptance/workflow.ts` execute all 19 cases using
`patchOffice()`. Every Then asserts receipt counts, source/destination custody,
reopened values, style closure or preservation as specified. Extra unit tests
cover concurrent guarded writers, fingerprints, path aliases, invalid XML and
overlapping targets. No inherited source-module/API parity rows are closed.

## Typed inputs and exchange version

Bun validation caught a single-backslash newline in v1's `value_json` data table:
Gherkin decoded it into a literal newline inside a JSON string, which JSON rejects.
The audit branch corrected that escape, added strict JSON compilation checks and
sealed v2. This checkout uses the sealed correction without local feature patches.
A regression verifies that one backslash fails and two preserve the JSON escape.
All four binary fixture hashes and 19 stable case identities remain unchanged.

The outcome interchange proposal is now version 2: it requires `stableCaseKey`
in addition to local case ID. Old version-1 records cannot silently satisfy the
new identity contract. `patchOffice()` is now a production export; a cross-language
outcome exporter remains separate work. The native acceptance ledger records the
execution feature hash and expanded steps; the sealed shared hash stays pinned.

## Fixture custody

The original 74 rcarmo archives and 1,681 frozen reference files are unchanged.
This pack adds four separately pinned derived fixtures; it does not redefine that
original denominator. Tests read exact checked-in pack bytes, with no fallback to
workspace exports or Python-generated inputs. Output paths are temporary and
separate from fixture paths. Package sentinels make opaque-part loss observable.

The pack includes Python preparation/validation scripts as reference material.
Bun verification never runs them. Upstream Python readback and deterministic
regeneration are historical evidence in the pack, not native Office rendering or
proof that the proposed mutation cases pass. Defective Python observations remain
diagnostics and must never become output goldens.

## Native workflow binding

Preview/strict/safe operations use native batch/staging code without MCP transport.
Receipts count only changed operations after atomic save succeeds. All modes
resolve every target first and refuse missing, ambiguous or overlapping targets.
The first safe mode also requires all targets; best-effort is not implemented. A distinct existing destination must survive refusal byte-for-byte; after
return no staging files may remain in the document directory.

For the cache case, select `invalidate-without-recalculation` even when a native
calculation engine later exists. Preserve the formula, clear the dependent cache,
and report recalculation required. Native bindings assert both receipt state and reopened formula/cache values.
Independent native Office calculation/rendering checks remain unverified.

Before publication, review fixture redistribution permission. The MCP source has
no root licence at the pinned revision; original template/library notices remain
inside the pack. See `pack/README.md` for the source adapter and preservation rules.
