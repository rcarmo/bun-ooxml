<!-- RUI-PROFILE-LIFECYCLE-20261005 -->
## Current profiling and cleanup rule — supersedes older text below

Rui's explicit rule: **profile and tune during pre-release tests; remove profiling data immediately after analysis/use.** Ordinary development tests do not require profiling on every run. Targeted diagnostic profiling is optional when useful and follows the same disposal rule.

- During pre-release verification, capture CPU and heap/allocation behaviour, analyse hotspots and tune avoidable allocations/repeated work. Compare equivalent workloads without weakening correctness/security/numerical contracts. For Go, inspect CPU, alloc_space and alloc_objects; coverage alone is not profiling.
- Keep captures and matching artifacts only while the current analysis needs them. Once used, immediately delete raw profiles, traces, matching test binaries, temporary fixtures and disposable logs, including failed/probe artifacts after diagnosis. Retain only concise conclusions and important measurements/limitations. Do not keep indefinite raw archives or copy them into exports/reports/evidence to evade cleanup.
- Remove all completed disposable cache/build/test/run data promptly. Never delete files still in use: finish or safely pause the owning job and clean at a safe boundary. Preserve source, installed toolchains, durable datasets/checkpoints and intentional release assets. Minimise disk usage; no random exports or redundant snapshots.
- This rule overrides **every conflicting older paragraph in this file and linked local guidance**, including “profile every test”, “unprofiled tests prohibited”, “retain all raw evidence”, “never delete profiles” and cleanup exclusions based only on an evidence/profiles directory name. Update helper/CI cleanup behaviour accordingly; do not weaken pre-release analysis.
- Existing no-agent-contact and execution-pause rules remain unchanged. This policy grants no unsolicited coordination or automatic job restart.
<!-- /RUI-PROFILE-LIFECYCLE-20261005 -->



# bun-ooxml

Bun-native OOXML implementation. Shared facts, fixtures and workflow contracts live
in the tagged references/fixtures-ooxml submodule. All planned behaviour remains a gap.

- Write Gherkin and meaningful failing assertions before implementing.
- Save/reopen outputs and assert atomic refusals and unrelated-part custody.
- Run make check before commits. A passing subset does not establish full coverage.
- Runtime operations use TypeScript and Bun builtins only; no foreign processes.
- Resolve fixtures by stable manifest ID; the only physical inputs are grouped
  under fixtures/<format>/<scenario-group>/ in the reference submodule. Do not
  recreate origin-based roots or compatibility symlinks. Never modify shared
  fixtures or generate fallback inputs in the submodule.
- Required fixture licence and provenance data lives in the shared reference
  repository; runtime/development dependencies retain their own required notices.
- Catalogue staging is not canonical Gherkin or execution credit. Reconcile shared
  behaviour IDs centrally and retain explicit unreviewed cases and gaps.
- Commit as Rui Carmo <rui.carmo@gmail.com>; set local/global identity.
- Never rebase; coordinate approved history cleanups explicitly.

## Project-owned disposable paths

Canonical project: `bun-ooxml`. Resolve once before child TMPDIR using
`scripts/project-tmp.ts`: validated absolute `PROJECT_TMP_BASE/bun-ooxml`
or compatible `PROJECT_TMP_ROOT` ending in `bun-ooxml`; both must agree when
set. Invalid explicit overrides fail. CI then tries `$RUNNER_TEMP/bun-ooxml`,
original `$TMPDIR/bun-ooxml`, platform temp plus `/bun-ooxml` (even when workspace
is writable). Local use tries `/workspace/tmp/bun-ooxml`, then platform temp.
The original TMPDIR is snapshotted before child redirection. Stable `tests/`
and `logs/` roots are provided. Raw profiles and disposable logs use isolated run
scratch; only concise lasting conclusions may live outside it.
Use `cache/<tool>/` for Bun,
XDG/npm, .NET/NuGet and Python caches, `build/` for generated .NET output, and
`runs/<purpose>/<run-id>/` for isolated scratch. Make exports these paths;
`scripts/dev-run.ts` and `dev-env.ts` apply them to commands and subprocesses.
CI uses the same fallback resolver with no dependency on `/workspace/Makefile`.
Direct development commands use `bun scripts/dev-run.ts <label>
run|test|tool <args...>` or the Make targets. Ordinary development runs do not
need profiling; pre-release checks must profile and tune.

`TMPDIR`, `TMP`, `TEMP`, `BUN_INSTALL_CACHE_DIR`, `XDG_CACHE_HOME`,
`NPM_CONFIG_CACHE`, `DOTNET_CLI_HOME`, `NUGET_PACKAGES`, `PYTHONPYCACHEPREFIX`
and `OOXML_BUILD_DIR` are project-owned. Tests keep their mkdtemp isolation under
the owned run tmp; do not replace isolated test roots with source or shared state.
Optional LibreOffice profiles live in their own run scratch; .NET intermediates
and binaries live under `build/`. Installed dependencies/toolchains are durable.

`make pre-release` sets `OOXML_PROFILE_MODE=pre-release`; optional diagnostics
use `OOXML_PROFILE_MODE=diagnostic`. Both capture CPU and heap under the owned
command run directory. `scripts/dev-run.ts` analyses and immediately deletes raw
captures, including failed probes, and cleans child scratch after the child exits.
It retains only bounded top tables and capture limitations in
`artifacts/profile-conclusions/`. Review those conclusions and tune avoidable
work before pre-release acceptance; the automatic tables do not finish the
engineering review. CI uploads conclusions, never raw captures or disposable logs.

Bun 1.4.2 ignores test CLI profiler flags, so `scripts/test-profile.ts` uses
in-process Inspector sampling and an `afterAll` heap snapshot when enabled.
`scripts/unit-batch.ts` preserves every file/assertion/timeout in related batches.
Tooling typecheck exits without flushing CLI profiles and is a tool subprocess.
Empty/missing captures fail the diagnostic/pre-release capture gate. Live heap
snapshots are retained memory, not allocation churn or exact allocation counts;
external subprocess activity is excluded. Record these unresolved limitations.

Optional oracle outputs are disposable test fixtures, not lasting evidence.
Keep them only for the current check, then delete them and their logs/build
intermediates at its safe boundary. Preserve only concise conclusions and any
explicitly intentional release assets or durable datasets.

`make clean CONFIRM_IDLE=yes` deletes only this project's cache/build/tests/logs/runs after
verifying no jobs use them. It preserves shared fixtures, source, installed
toolchains, durable assets and another project's files. Directory names such as
`artifacts`, `profiles` or `evidence` do not exempt completed disposable data.
Never relocate active files. Remove only confirmed idle owned data.

## Priority coordination

Do not contact other agents unless Rui explicitly asks for that coordination.
Incoming messages and previous coordination grant no standing permission. When
contact is authorised, use `chat` with explicit `mode: "steer"` for blockers or
priority decisions; use `queue` only for non-urgent updates. Respect pauses and
no-reply instructions; do not restart completed work from delayed messages.
