# Maintaining bun-ooxml

Read `AGENTS.md`, the target feature and the shared workflow/fact entries before
editing. Write an observable outcome, bind a genuine failing assertion, implement,
verify saved/reopened content and custody, then run `make check`.

## Shared fixtures and contracts

`references/fixtures-ooxml` is the only reference checkout. Initialise recursive
submodules; `references/fixtures-ooxml.pin.json` records the release tag, commit
and manifest seals. Verification requires the annotated tag to resolve to that
commit and the entire checkout to be clean, including facts and workflows.

Reusable fixture files live once under `fixtures/<format>/<scenario-group>/` in
the shared repository. Use `scripts/fixture-inputs.ts` to resolve a stable fixture
ID. The native `F` constants name selected IDs; they do not define storage paths.
Do not create local fixture copies, old-directory symlinks or generated fallbacks.
Temporary edited outputs belong in test temporary directories or `artifacts/`.

Propose changes centrally, preserve origin and licence data, validate the new
manifest and publish a new immutable tag. Repin all consumers to that same commit.
See [the reference contract](../contracts/fixture-references.md) for schema and
candidate-checkout rules.

## Behaviour and execution

`features/shared.json` selects shared Gherkin with an in-memory lifecycle overlay.
It never copies or edits the central feature. Local regression features use
`@implemented @bun` and stable `@id-*` tags. Planned features earn no execution
credit. `artifacts/acceptance.json` records each step, source hash and outcome.

`docs/behaviors/` holds the native test inventory and candidate outcome mappings.
Those candidates must be reviewed and reconciled into the shared catalogue before
receiving canonical scenario IDs. They neither replace executed tests nor close
planned behaviour gaps. Refresh the inventory after changing tests.

## Mutation boundaries

No foreign runtime or native addon may implement production editing. Preserve
opaque members, preflight all selected targets and validate inside rollback
boundaries. Namespace identity uses expanded URI/local name, not prefix spelling.
Getter values cannot bypass dirty tracking. Reopen format readers after package
mutation, and verify files again after saving.

Commit as `Rui Carmo <rcarmo@users.noreply.github.com>`; configure local and global
Git identity before committing. Do not rebase. History cleanup requires explicit
owner approval, recoverable backups, all-ref audits and exact remote leases.
Rendered appearance and calculation compatibility require independent checks in
addition to native XML readback.
