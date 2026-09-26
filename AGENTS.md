# bun-ooxml

Bun-native OOXML implementation. Shared facts, fixtures and workflow contracts live
in the tagged references/fixtures-ooxml submodule. All planned behaviour remains a gap.

- Write Gherkin and meaningful failing assertions before implementing.
- Save/reopen outputs and assert atomic refusals and unrelated-part custody.
- Run make check before commits. A passing subset does not establish full coverage.
- Runtime operations use TypeScript and Bun builtins only; no foreign processes.
- Never modify shared fixtures or generate fallback inputs in the submodule.
- Required licence and provenance data lives in the shared reference repository.
- Commit as Rui Carmo <rcarmo@users.noreply.github.com>; set local/global identity.
- Never rebase; coordinate approved history cleanups explicitly.
