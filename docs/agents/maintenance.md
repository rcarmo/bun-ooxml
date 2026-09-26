# Maintaining bun-ooxml with agents

Read `AGENTS.md`, `docs/contracts/port-scope.md`, the target feature and its current
source mapping before editing. Full parity includes inherited code and tests;
OOXML additions alone cannot close a format.

## Change loop

Write or refine a Gherkin outcome first. Bind it to an executable test and record
a genuine behaviour failure. Implement the smallest complete operation, verify
saved/reopened output, then run one relevant unit/type batch and commit it. Run
`make check` at integration boundaries. Avoid repeatedly running the full fixture
corpus while changing a single helper; the final batch still covers it.

Each implemented feature has `@implemented @bun`; each scenario has one `@id-*`.
Planned features live under `features/planned/` with `@planned`. A planned scenario
is backlog, never a skipped green test. Background and outline steps are compiled
with the official Gherkin implementation. Undefined, ambiguous, missing and failed
steps fail acceptance. `artifacts/acceptance.json` is replaced for every run and
contains hashes, expanded cases, source locations and per-step outcomes.

Assertions need observable effects: exact reopened text, valid relationship targets,
cell values/styles/caches, untouched payload hashes, or typed refusal plus unchanged
bytes. Counting scenarios or finding an assertion in a file cannot prove semantics.
Human/independent agent review checks whether each Then is actually established.

## Source custody

`references/manifest.json` pins revisions and every imported file SHA256. Both rcarmo
fixture corpora contain 37 Office archives; run integrity checks before and after
execution. Never rewrite fixtures from tests. Use in-memory copies or temporary
output directories. A source update is an explicit reviewed import, not a fixture fix.

`docs/contracts/parity-ledger.json` starts at source/test file granularity. Expand
each row to API/test case and Gherkin step mappings as it is ported. A hash match
proves provenance only. Set `mapped` only after reviewed assertions cover the whole
row, with all partial/refused/manual cases recorded. The full-parity gate also
requires no planned features. Do not reduce the denominator to improve a percentage.

## Shared Python audit contracts

Read `docs/contracts/shared-v2/README.md` before binding shared mutation workflows.
The eight format-specific scenarios expand to 19 planned cases with four pinned
derived inputs. Historical `@id-office-*` IDs are archived aliases, never extra
coverage. Historical defect observations are diagnostics, never golden outputs.
Preserve native-runner versus MCP-transport identity, stable case keys, typed JSON
values and actual input hashes. `make check` verifies the separate pack/audit
manifests and ID migrations; it does not execute Python scripts.

## Implementation invariants

* Preserve opaque members; parse only what the operation owns. Getter results must
  not allow callers to bypass dirty tracking. Reopen after save.
* Validate all targets before mutating a multi-target operation. Typed refusal must
  leave model state and disk bytes unchanged. Tests need at least one valid target
  before a failing one to detect partial writes.
* Keep text offsets bound to the inspected XML. Escaped user text and replacement
  XML fragments are different inputs. Never concatenate unescaped text into XML.
* A cell edit can invalidate caches on other sheets. A style index is valid only
  when its style-table dependency exists in the output.
* Use Bun builtins for I/O, hashes, compression and tests. Bun's `node:zlib` and
  `node:fs` imports execute Bun's built-in implementations; they do not launch Node.
  `inflateRawSync` supplies a hard output cap absent from the simple inflate API.
* Production imports cannot launch foreign runtimes, converters or native addons.
  Frozen Python files are read-only reference code, never execution dependencies.

## File ownership

`src/opc/` owns archive and relationship custody; `src/xml/` owns strict parsing and
offsets; format directories own semantic edits. `tests/acceptance/steps.ts` combines
bindings. Keep format-specific step phrases distinct to avoid ambiguous bindings.
Parallel agents need explicit file ownership, then one integration/type/acceptance
batch. Only the integrating agent commits shared work.

## Current hazards

The first slices are deliberately smaller than the full source contract. All three
formats now use shared OPC custody and atomic path saves. DOCX also returns output
bytes; XLSX omitted save paths overwrite the opened path. Examples use explicit
output paths. Native
spreadsheet calculation and producer/rendering comparisons have not been implemented.
These differences need explicit tests and mappings before a stable public release.
