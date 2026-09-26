# bun-ooxml

Bun-native TypeScript for editing existing Word documents, PowerPoint decks and
Excel workbooks. I want document edits to stay inside the Bun process while
preserving the Office package parts an edit did not touch.

The first slice handles guarded cross-run text replacement, slide text and notes,
and existing-cell edits with formula-cache invalidation. ZIP, XML and OPC code
use Bun's built-in file, hash and compression implementations. There are no
runtime package dependencies or Office subprocesses.

Full ports of OOXML DOCX, OOXML PPTX and OOXML XLSX, including their inherited
APIs, are in progress. This version does not implement document creation,
tracked revisions, slide imports, structural spreadsheet edits or calculation.
The [scope and closure criteria](docs/contracts/port-scope.md) and
[behaviour ledger](docs/contracts/parity-ledger.json) track that work.

```ts
import { Document } from "bun-ooxml";

const document = await Document.open("contract.docx");
const matches = document.find("thirty calendar days");
if (matches.length !== 1) throw new Error("Expected exactly one target");
matches[0]!.replace("thirty business days");
await document.save("contract-edited.docx");
```

Read [agent usage](docs/agents/usage.md) before editing unfamiliar files. It covers
refusals, target lifetimes and save/reopen checks. [Agent maintenance](docs/agents/maintenance.md)
explains source pins and the Gherkin-to-test-to-outcome workflow.

## Development

Requires Bun 1.4.1 or newer. `make install` installs development tools; `make check`
runs types, frozen-source checks, unit tests, exact Gherkin acceptance and examples.
`make parity` also requires all planned work and source/test gaps to be closed.
It currently fails as expected.

Tests use committed fixtures from `rcarmo/go-ooxml` and
`rcarmo/python-office-mcp-server`. Frozen Python sources under `references/` are
reference material; the Bun library never imports or executes them.

MIT. See [third-party notices](THIRD_PARTY_NOTICES.md) for source and fixture provenance.
