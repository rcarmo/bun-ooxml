# bun-ooxml

Bun-native TypeScript for creating and editing Word documents, PowerPoint decks and
Excel workbooks. Edits run inside the Bun process and preserve untouched Office
package parts.

Authoring supports DOCX paragraphs/tables, PPTX title slides/tables and XLSX sheets/cells
entirely in Bun. Slides also accept [positioned text boxes](docs/contracts/text-box.md)
with explicit geometry and optional direct bold/italic flags.
[Effective bold/italic inspection](docs/contracts/effective-formatting.md) resolves
a bounded paragraph-style cascade with provenance and explicit unsupported contexts.
[Slide reordering](docs/contracts/slide-order.md) preserves slide identities and
package parts. Existing-file edits include guarded cross-run replacement, slide
text, [existing speaker notes](docs/contracts/notes-editing.md) and cell edits
with bounded formula-cache invalidation. Existing XLSX
[cell-style selection](docs/contracts/cell-style.md) preserves values, formulas and caches. Word paragraph
runs support direct bold/italic overrides, [half-point font sizes](docs/contracts/font-size.md)
and existing paragraph-style selection without text or style-graph changes. Bounded [style authoring](docs/contracts/style-authoring.md)
adds named paragraph styles with validated base chains and direct bold/italic flags.
[Page geometry](docs/contracts/page-layout.md) changes the final Word section's
size and margins while preserving earlier sections and text.
`patchOffice()` adds
read-only previews, all-targets-required batches, guarded saves and per-target
receipts. It can track one unique Word replacement with explicit author/date and
report saved revision IDs separately from preview counts. Guarded OPC graph
helpers add/detach parts and relationships; package diffs identify payload and
content-type changes. Package-level Word review APIs inspect linked stories,
resolve supported run-level revisions and author bounded tracked replacements.
Existing comment threads can be inspected and individual resolution flags changed
without touching comment bodies or anchors.
[`admitPackage()`](docs/contracts/package-admission.md) validates bounded ZIP
structure and XML members without requiring a complete OPC graph.
[`xmlEquivalent()`](docs/contracts/xml-comparison.md) provides conservative
XML comparison for the documented preservation profile.
[`XmlSnapshot`](docs/contracts/xml-removal.md) removes disjoint XML subtrees
without rewriting surrounding source characters.
[`comparePackageArchives()`](docs/contracts/package-comparison.md) separates
those equivalent XML payloads from binary changes, additions and removals.
ZIP, XML and OPC code use Bun's built-in file, hash and compression implementations. There are no
runtime package dependencies or Office subprocesses.

General Word Compare, move/format/table revisions, comment authoring, slide imports,
merged-table restructuring, chart authoring, structural spreadsheet edits and
formula calculation are unsupported. [Direct A1 range parsing](docs/contracts/a1-ranges.md)
is available independently; it does not edit references or evaluate formulas.
See [supported operations and limits](docs/contracts/port-scope.md).

```ts
import { Document } from "bun-ooxml";

const document = await Document.open("contract.docx");
const matches = document.find("thirty calendar days");
if (matches.length !== 1) throw new Error("Expected exactly one target");
matches[0]!.replace("thirty business days");
await document.save("contract-edited.docx");
```

Use the [native workflow API](docs/contracts/workflow-api.md) for preview/strict/safe
batches. The [usage guide](docs/agents/usage.md) covers refusals, target lifetimes
and save/reopen checks.

## Specification

The shared [`fixtures-ooxml` specification][fixtures] defines expected behaviour,
MIME types, namespaces and relationship facts. Its workflows and fixtures provide
the reference for validation. Fixtures are stored under
`fixtures/<format>/<scenario-group>/` and resolved by stable manifest ID. See
[shared references](docs/contracts/fixture-references.md).

## Development

Requires Bun 1.4.1 or newer. `make install` installs development tools; `make check`
runs type checks, test-inventory and mapping checks, shared-reference validation,
unit tests, Gherkin acceptance tests and examples. See the
[test catalogue](docs/behaviors/README.md) for details. `make parity` also requires
all specified behaviour to be implemented; it currently fails because some
operations are unsupported.

`make office-oracles` separately checks three authored files with Open XML SDK,
LibreOffice and Poppler. It catches schema errors, missing PDF text and one
external calculation round trip. See [independent checks](docs/contracts/office-oracles.md)
for prerequisites and limits; these tools are never production dependencies.
`make property-campaign` runs fixed-seed XML/ZIP/Office properties and five bounded
timing workloads. [Campaign scope](docs/contracts/property-campaign.md) describes
replay inputs, refusal accounting and measurement limits.

Clone with `--recurse-submodules` or run `git submodule update --init --recursive`
to obtain the shared specification and fixtures. Fixtures are read-only; missing
inputs fail rather than generating fallbacks.

MIT. See [third-party notices](THIRD_PARTY_NOTICES.md) for licences and fixture provenance.

[fixtures]: https://github.com/rcarmo/fixtures-ooxml
