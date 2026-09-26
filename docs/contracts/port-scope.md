# Behaviour scope

The target covers DOCX/PPTX/XLSX package preservation, authoring, formatting,
comments/revisions, composition, charts, structural spreadsheet editing and
calculation. Current bounded slices do not close this scope.

Shared workflow contracts, expected outcomes, facts and fixture provenance are in
references/fixtures-ooxml. Consumer-only regression Gherkin and bindings stay local.
All planned contracts remain gaps; replacing external-source inventories does not
turn unimplemented behaviour into completed work.

Production runs inside Bun. External Office applications may be independent test
oracles only. XML readback does not establish rendering or calculation fidelity.
