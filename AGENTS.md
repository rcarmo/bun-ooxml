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
- Commit as Rui Carmo <rcarmo@users.noreply.github.com>; set local/global identity.
- Never rebase; coordinate approved history cleanups explicitly.

## Priority coordination

Use `chat` with `target_agent_name: "@alias"` and explicit `mode: "steer"` for
scope changes, stop/hold requests, release corrections, safety blockers and
unblocking decisions. Reserve `mode: "queue"` for routine progress. Name the
current revision, requested action, owner and superseded notice. Receivers verify
current state and acknowledge once; delayed messages must not restart obsolete
work. Use `session_control` only for runtime/session operations.
