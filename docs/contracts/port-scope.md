# Port scope

The target is full behavioural ports of OOXML DOCX 0.2.1, OOXML PPTX 0.2.1 and
OOXML XLSX 0.2.2, including inherited OOXML, OOXML and OOXML APIs,
package custody, typed refusals, comments/revisions, composition, charts,
formatting, structural edits and formula-cache contracts. This is a multi-slice
implementation; a small passing core does not close that scope.

The public API uses TypeScript naming and async file I/O. Python's import names,
object protocols and implementation details need explicit mappings, not literal
syntax translation. Each source module and inherited test belongs in the parity
inventory; API/test-level mappings refine that inventory as slices are ported.

The Python MCP server supplies fixtures and workflow requirements. A new MCP
transport is not required to make the core library usable. Its tools are inventoried
separately so they cannot silently replace the wider library contract.

## Native execution

Production code uses Bun file, compression, hashing and runtime primitives with
TypeScript OOXML/XML logic. No foreign runtime, Office subprocess or remote API
may perform document operations. LibreOffice-backed oracle behaviour is a known
architecture decision: the target is native calculation/certification semantics;
LibreOffice equivalence cannot be asserted without a separate comparison corpus.
External Office applications may be used as test oracles, never runtime dependencies.

## Frozen inputs

Reference repositories are read-only. Fixture origin, revision and SHA-256 are
recorded for both rcarmo/go-ooxml and rcarmo/python-office-mcp-server. Uncommitted
MCP changes are excluded. OOXML source and tests are pinned, including inherited
surfaces; OOXML additions alone are insufficient for full parity.

## Cross-language limits

`cross-language-limits.md` records bounded revision/cache/transport behaviour and
an unresolved extended-comment content-type discrepancy. Its planned scenarios
must be closed by independent evidence; shared workflow passes do not cover them.

## Closure

Every imported behaviour has a reviewed mapping to Gherkin, executed steps and
saved/reopened outcomes. No missing, skipped, undefined, stale or empty results
pass acceptance. All planned behaviours and per-source test gaps must be closed
before the full-parity gate passes. Visual layout and spreadsheet computation
require independent producer/calculation checks in addition to XML assertions.
