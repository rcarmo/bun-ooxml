# Project cache and temporary paths

Bun development commands now resolve one project-owned disposable root before
changing child environment variables. Runtime APIs, shared specifications,
fixtures and released reference pins are unchanged.

## Resolution and cleanup

`scripts/project-tmp.ts` selects explicit `PROJECT_TMP_BASE/bun-ooxml` or a
compatible `PROJECT_TMP_ROOT`; both must agree when set. Invalid empty, relative,
conflicting or symlinked explicit paths fail. CI tries `RUNNER_TEMP`, original
`TMPDIR`, then the platform temporary directory, each with a `bun-ooxml` child.
Local use tries writable `/workspace/tmp/bun-ooxml`, then platform temp.
The stable project tree contains `cache`, `build`, `tests`, `logs` and `runs`.

Make, package scripts and the development launcher route Bun/XDG/npm/.NET/NuGet/
Python caches and TMPDIR/TMP/TEMP through that tree. Each test process owns its
isolated run scratch. The resolved physical scratch path avoids Bun dynamic-import
failures through this host's `/workspace` alias. LibreOffice profiles live in
run scratch; .NET intermediates/binaries live under `build/dotnet/schema`.
Legacy source-local `obj`/`bin` are excluded from .NET source discovery, without
removing old files or moving active jobs.

`make clean CONFIRM_IDLE=yes` removes only the resolved project's disposable
subtrees. It does not remove `artifacts`, profiles, fixtures, dependencies or
source. Optional oracle outputs use new `artifacts/policy-oracles/<run-id>`
directories. Existing retained output stays in place; no legacy cleanup ran.

## Verification

- Type checking and native inventory/mapping checks pass; declaration identities
  and mapping/execution credit are unchanged. Only the policy test-support hash
  was added to the local inventory.
- All 177 unit files execute once in related batches: 1539 passing tests including
  three new policy controls, 52 existing skips, no failures. No assertion or test
  timeout was relaxed. Interrupted batches were resumed from retained receipts.
- Acceptance passes 752/752 implemented cases; 59 planned cases remain visible.
  Examples, reference verification and shared-contract inventory checks pass.
- Explicit override, BASE/ROOT conflict, CI-vs-workspace precedence, original
  TMPDIR fallback, traversal and symlink controls pass. CI-style installation
  uses the resolver without the workspace Makefile.
- Relocated .NET restore/build and the existing Office oracle pass schema,
  rendering, revision/style controls and arithmetic roundtrip. Graphics and
  sealed-source SmartArt UNO checks pass with project-owned LibreOffice profiles.
- The clean graphics candidate is explicitly retained at shared `0cf83c1` under
  `/workspace/tmp/fixtures-ooxml/runs/pinned/shared-v043-mDKpt4`. Shared host-policy
  publication does not advance this consumer's reference pin.

## Profiling analysis

Each test batch captures process-local Inspector CPU samples and a Chrome live
heap snapshot. Bun 1.4.2 ignores CLI test profiler flags; the initial missing
captures were failures. In-process JSC raw traces retained roughly 319 MB of CPU
trace data and 315 MB of heap, distorting full-suite timings. Compact Inspector
capture reduced a comparable full run to roughly 32 MB CPU and 46 MB heap.
Related test batches bound catalogue AST retention without weakening assertions.

Sampling is recorded per command: 10 ms normally, 100 ms for batches containing
expensive comment-thread/move probes (including the runtime batch), 1 ms for the
short policy/oracle-runner batch. Representative runtime and acceptance profiles contain 510 and 5127
CPU samples respectively. This is capture verification, with no throughput or
allocation-reduction claim; host load and instrumentation make unlike-run timing
comparisons unsuitable.

Catalogue CPU is dominated by Babel parsing and `inventoryNativeTests` (test
harness). Runtime CPU centres on ZIP writing, XML scanning and content-type
readback. Acceptance likewise spends time in XML/ZIP setup and shared workflow
execution. Live heaps include parser/function code, small object/string records
and an 8 MB adversarial test string. These are unresolved profiling observations;
no runtime optimisation belongs to this path-only change. The policy batch's CPU
centres on subprocess setup/kill and filesystem isolation. Git, .NET,
LibreOffice and child Bun oracle commands are excluded from parent profiles;
Bun snapshots measure retained live memory, not allocation churn or Go
alloc_space/alloc_objects.

Raw commands, logs, profiles and analyses remain in `artifacts/policy-profiles`.
`policy-batched-final.log`, `policy-resume.log` and `policy-finish.log` record the
completed/resumed gate. Failed captures, interrupted runs and timeout probes are
retained separately and do not count as successful verification.
