# Maintaining bun-ooxml

Read AGENTS.md, the target feature and shared workflow/fact entries before editing.
Write an observable outcome, bind a genuine failing assertion, implement, verify
saved/reopened content and custody, then run make check.

Shared inputs live in references/fixtures-ooxml. Initialise recursive submodules.
The fixtures symlink is compatibility navigation, not a second corpus. Never write
to shared inputs. Missing or mismatched pins must fail before execution. Propose
new reference data centrally, publish a new immutable tag and repin all consumers.

features/shared.json selects shared Gherkin with an in-memory lifecycle overlay.
It never copies or edits the central feature. Local implemented features use
@implemented @bun and stable @id-* tags; planned features earn no credit.
artifacts/acceptance.json records each executed step, source hash and case outcome.

No foreign runtime or native addon may implement production editing. Preserve
opaque members, preflight all selected targets and validate inside rollback
boundaries. Namespace identity uses expanded URI/local name, not prefix spelling.
Getter values cannot bypass dirty tracking. Reopen after mutation/save.

Commit as Rui Carmo <rcarmo@users.noreply.github.com>; configure local/global Git
identity. Shared contract inventories never substitute for local execution evidence.
Full rendering and calculation compatibility still require independent checks.
