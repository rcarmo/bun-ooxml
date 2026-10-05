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
subtrees, preserving fixtures, installed dependencies, durable assets and
source. Directory names such as `artifacts` or `profiles` do not exempt completed
disposable output. Optional oracle fixtures last only for the current check.

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

During the earlier policy verification, each test batch captured process-local
Inspector CPU samples and a Chrome live heap snapshot. The corrected lifecycle
profiles pre-release tests and optional diagnostics; ordinary development tests
can run without profiling. Bun 1.4.2 ignores CLI test profiler flags; the initial missing
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

After the 2026-10-05 lifecycle correction, completed raw profiles, failed probe
captures, disposable logs, temporary oracle fixtures and rebuildable outputs
were deleted at an idle boundary. This note retains the conclusions and key
measurements. No raw archive or copy was kept elsewhere.

`scripts/dev-run.ts` now writes diagnostic/pre-release captures under the owned
run directory, analyses them, and immediately deletes raw data. Only concise
conclusions remain in `artifacts/profile-conclusions/`. CI excludes raw data from
uploads. Capture failures remain explicit; conclusions still require engineering
review and tuning before pre-release acceptance.

Lifecycle regression verification passed the ordinary development gate and the
20-test related policy/Gherkin/process batch in both ordinary and diagnostic
modes. The diagnostic run captured 38 CPU samples; most cumulative time was
asynchronous process/timeout handling, with live heap led by roughly 1.33 MB of
function code and 1.31 MB of byte arrays. This focused workload does not measure
runtime throughput or allocation churn. Raw captures were deleted immediately.
Missing capture, failed command, bounded conclusions and owned-run cleanup have
explicit controls. More than 2.1 GB of completed raw policy captures plus idle
cache/build/clone scratch were removed; source and installed dependencies stayed
in place.
