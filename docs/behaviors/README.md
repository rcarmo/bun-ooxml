# Native behaviour catalogue staging

The AST inventory contains 227 test declarations in 29 native test files. Five
candidate mapping files cover those declaration IDs once each. Semantic review
and reconciliation into the shared Gherkin catalogue are unfinished. A bounded
review of the 18 XML/QName declarations is recorded in the shared
`ledgers/consumers/bun-xml.json`, with partial and unmapped assertions explicit.
The shared parsing/QName features now replace their former local copies.

| Candidate group | Declarations |
|---|---:|
| Package, ZIP and graph | 65 |
| Word documents and reviews | 45 |
| Slides, spreadsheets and cache boundaries | 40 |
| XML, namespace and format guards | 38 |
| Workflows, references and acceptance runner | 39 |

`native-test-inventory.json` records declaration identity, suite/title, source
location/hash, assertion expressions and loops. It counts declarations, not
runtime-expanded leaves. Twenty-five declarations contain parameters or loops
that need explicit expansion review. Assertions hidden in helper calls also need
manual inspection; an empty extracted assertion list does not mean no assertion
ran.

The `staging-*.json` files contain proposed preconditions, operations and outcomes.
Their `reviewState` is `candidate-needs-parent-review`; they have no canonical
scenario IDs and grant no new Gherkin execution credit. They were drafted before
the grouped fixture migration. Review fixture/path and inventory descriptions
against current tests before importing them centrally, even when test IDs match.

Refresh the native inventory with:

```sh
bun run scripts/test-inventory.ts
```

For each candidate, inspect the current assertion and helper behaviour, expand
parameter variants, and record any weak, conditional or missing checks. Reuse the
central scenario ID for equivalent behaviour. Conflicting preconditions or
outcomes need a separate scenario or an explicit issue, not contradictory copies.

Only reviewed functional Gherkin enters `fixtures-ooxml`. Per-language ledgers
retain test-to-scenario mappings and execution results. Staging descriptions remain review inputs. The shared XML mapping supersedes
XML-specific candidate prose here; remaining non-XML rows still need review. The existing 114 executable Bun cases remain the current acceptance
result; candidate mapping counts do not change that result.
