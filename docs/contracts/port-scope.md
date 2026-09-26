# Behaviour scope

The target covers DOCX/PPTX/XLSX package preservation, authoring, formatting,
comments/revisions, composition, charts, structural spreadsheet editing and
calculation. Current bounded slices do not close this scope.

Shared workflow contracts, expected outcomes, facts and fixture provenance are in
references/fixtures-ooxml. All implemented Gherkin is selected from that checkout;
consumer bindings, native regressions and planned backlog features stay local.
All planned contracts remain gaps; replacing external-source inventories does not
turn unimplemented behaviour into completed work. The complete native test
catalogue is being captured and reconciled; staging mappings are not coverage.

Production runs inside Bun. External Office applications may be independent test
oracles only. The [authored-file oracle lane](office-oracles.md) checks three
small samples with Open XML SDK schema validation, LibreOffice PDF text/page
checks and one external formula recalculation. Broader rendering/calculation
fidelity and full-corpus validation remain open.
