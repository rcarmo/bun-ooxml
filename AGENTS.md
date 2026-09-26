# bun-ooxml

One repository for full Bun-native TypeScript ports of OOXML, OOXML,
OOXML and their inherited OOXML behaviours. Full parity is the target;
unfinished behaviours stay explicitly planned in the ledger.

## Workflow

* Read `docs/contracts/port-scope.md` and the relevant feature before editing.
* Write Gherkin first, then executable assertions, record a genuine behavioural
  failure, implement, and verify the saved/reopened observable outcome.
* Follow rcarmo/minicore lifecycle tags and rcarmo/gi source-pinned per-step
  ledgers. Scenario counts alone do not establish behaviour.
* `@implemented @bun` requires every expanded scenario and step to pass.
  `@planned` is backlog, never a skip or a pass. No placeholder implementations
  count as a full port. Unsupported upstream behaviour and unported behaviour
  are different states.
* Pin upstream source, inherited tests and fixture bytes. Never edit frozen
  sources or expected results to make the port pass.
* All runtime code is TypeScript executed by Bun. No Python, Go, Node process,
  subprocess Office converter, FFI/native addon, hosted service or MCP backend
  may implement a runtime operation. Bun-provided builtins are allowed.
* Run `make check` before each commit; keep units, acceptance and full-parity
  gates distinct. A passing implemented subset is not full parity.
* Preserve original bytes for untouched parts. Refuse ambiguous or unsafe edits
  before mutation. Bound untrusted input resources. Test failed edits for rollback.
* Work only in assigned files when tasks run in parallel. No rebase.
* Commit as Rui Carmo <rui.carmo@gmail.com>; configure local and global identity.
