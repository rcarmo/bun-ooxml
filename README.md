# bun-ooxml

Bun-native TypeScript for editing existing Word documents, PowerPoint decks and
Excel workbooks. I want document edits to stay inside the Bun process while
preserving the Office package parts an edit did not touch.

Initial authoring creates DOCX paragraphs/tables, PPTX title slides/tables and XLSX sheets/cells
entirely in Bun. Slides also accept [positioned text boxes](docs/contracts/text-box.md)
with explicit geometry and optional direct bold/italic flags. Existing-file edits include guarded cross-run replacement, slide
text/notes and cell edits with bounded formula-cache invalidation. Existing XLSX
[cell-style selection](docs/contracts/cell-style.md) preserves values, formulas and caches. Word paragraph
runs support direct bold/italic overrides and existing paragraph-style selection
without text or style-graph changes. Bounded [style authoring](docs/contracts/style-authoring.md)
adds named paragraph styles with validated base chains and direct bold/italic flags.
`patchOffice()` adds
read-only previews, all-targets-required batches, guarded saves and per-target
receipts. It can track one unique Word replacement with explicit author/date and
report saved revision IDs separately from preview counts. Its original native
bindings pass the 19 shared mutation cases. Guarded OPC graph
helpers add/detach parts and relationships; package diffs identify payload and
content-type changes. Package-level Word review APIs inspect linked stories,
resolve supported run-level revisions and author bounded tracked replacements.
Existing comment threads can be inspected and individual resolution flags changed
without touching comment bodies or anchors.
ZIP, XML and OPC code
use Bun's built-in file, hash and compression implementations. There are no
runtime package dependencies or Office subprocesses.

Full DOCX, PPTX and XLSX behaviour coverage is in progress. Creation is limited to those initial paragraph/slide/cell surfaces. This version
does not implement general Word Compare, move/format/table revisions, comment authoring,
slide imports, merged-table restructuring or chart authoring,
structural spreadsheet edits or calculation.
The [scope and closure criteria](docs/contracts/port-scope.md) and
[shared workflow ledger](references/fixtures-ooxml/ledgers/workflows.json) track that work.

```ts
import { Document } from "bun-ooxml";

const document = await Document.open("contract.docx");
const matches = document.find("thirty calendar days");
if (matches.length !== 1) throw new Error("Expected exactly one target");
matches[0]!.replace("thirty business days");
await document.save("contract-edited.docx");
```

Use the [native workflow API](docs/contracts/workflow-api.md) for preview/strict/safe
batches. Read [agent usage](docs/agents/usage.md) before editing unfamiliar files. It covers
refusals, target lifetimes and save/reopen checks. [Agent maintenance](docs/agents/maintenance.md)
explains reference pins and the Gherkin-to-test-to-outcome workflow.

## Shared reference

Fixtures, MIME types, namespace and relationship facts, workflow Gherkin and
consumer mappings come from one tagged [`fixtures-ooxml` checkout][fixtures].
Fixtures are stored once under `fixtures/<format>/<scenario-group>/` there;
Bun tests resolve stable manifest IDs instead of copying files or rebuilding
producer-specific folders. See the [reference contract](docs/contracts/fixture-references.md).

The native test inventory and candidate behaviour mappings in `docs/behaviors/`
are review inputs for the common catalogue. They do not count as implemented
Gherkin or complete format coverage.

## Development

Requires Bun 1.4.1 or newer. `make install` installs development tools; `make check`
runs types, shared-reference checks, unit tests, exact Gherkin acceptance and examples.
`make parity` also requires all planned behaviour gaps to be closed.
It currently fails as expected.

Tests use the tagged shared reference submodule. Clone with --recurse-submodules
or run git submodule update --init --recursive. Fixtures are read-only; missing
inputs fail rather than generating fallbacks.

MIT. See [third-party notices](THIRD_PARTY_NOTICES.md) for licences and fixture provenance.

[fixtures]: https://github.com/rcarmo/fixtures-ooxml
