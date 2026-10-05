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
and `logs/` roots are provided; retained evidence stays outside disposable logs.
Use `cache/<tool>/` for Bun,
XDG/npm, .NET/NuGet and Python caches, `build/` for generated .NET output, and
`runs/<purpose>/<run-id>/` for isolated scratch. Make exports these paths;
`scripts/dev-run.ts` and `dev-env.ts` apply them to commands and subprocesses.
CI uses the same fallback resolver with no dependency on `/workspace/Makefile`.
Direct development commands use `bun scripts/dev-run.ts <label>
run|test|tool <args...>` or the Make targets; raw test execution is prohibited.

`TMPDIR`, `TMP`, `TEMP`, `BUN_INSTALL_CACHE_DIR`, `XDG_CACHE_HOME`,
`NPM_CONFIG_CACHE`, `DOTNET_CLI_HOME`, `NUGET_PACKAGES`, `PYTHONPYCACHEPREFIX`
and `OOXML_BUILD_DIR` are project-owned. Tests keep their mkdtemp isolation under
the owned run tmp; do not replace isolated test roots with source or shared state.
Optional LibreOffice profiles live in their own run scratch; .NET intermediates
and binaries live under `build/`. Installed dependencies/toolchains are durable.

Retain CPU/heap profiles, logs, receipts and generated datasets under
`artifacts/policy-profiles/` (or an explicit `OOXML_EVIDENCE_ROOT` outside the
project disposable root); existing graphics/acceptance evidence stays in place.
Optional oracle outputs use a new `artifacts/policy-oracles/<run-id>/` by default,
so they do not overwrite historical graphics evidence.
Every Bun test/check/campaign captures CPU and live heap profiles. Bun 1.4.2
ignores test CLI profiler flags, so `scripts/test-profile.ts` uses in-process
Inspector sampling and an `afterAll` heap snapshot. JSC's raw trace profiler
inflates full-suite memory/time, so the default uses compact Inspector profiles.
`scripts/unit-batch.ts` runs every unit file once in related catalogue/runtime/
policy batches, preserving original assertions and test timeouts. This isolates
retained catalogue ASTs without raising performance or correctness budgets.
Tooling typecheck may exit without flushing CLI profiles and is recorded as a
tool subprocess, not a test. Review cumulative hotspots after each run; automatic
top tables alone do not close the profiling gate. Record empty/missing capture
and external subprocess exclusions. Bun live heap bytes are not total allocation
churn or exact allocation counts. Go alloc_space/objects requirements apply to
Go owners separately.

`make clean CONFIRM_IDLE=yes` deletes only this project's cache/build/tests/logs/runs after
verifying no jobs use them. It never deletes artifacts, shared fixtures, source,
toolchains or another project's files. Never relocate active files. Old disposable
paths may be removed only after confirming they are idle; no broad legacy cleanup.

## Priority coordination

Use `chat` with `target_agent_name: "@alias"` and explicit `mode: "steer"` for
scope changes, stop/hold requests, release corrections, safety blockers and
unblocking decisions. Reserve `mode: "queue"` for routine progress. Name the
current revision, requested action, owner and superseded notice. Receivers verify
current state and acknowledge once; delayed messages must not restart obsolete
work. Use `session_control` only for runtime/session operations.
