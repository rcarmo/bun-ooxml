# Behaviour scope

bun-ooxml supports selected DOCX, PPTX and XLSX authoring and editing operations
while preserving unrelated package parts. The API contracts describe supported
targets and explicit refusals. General document comparison, comment authoring,
slide imports, chart authoring, merged-table restructuring, structural spreadsheet
editing and formula calculation are unsupported.

The shared specification in `references/fixtures-ooxml` defines expected outcomes,
format facts and fixture provenance. Bun acceptance tests select workflows from
that specification; unsupported operations are not enabled by their presence there.

Production operations run inside Bun. [Independent validation](office-oracles.md)
uses Open XML SDK, LibreOffice and Poppler on three generated samples. It checks
schema, PDF text/page properties and one external formula recalculation. General
Microsoft Office rendering compatibility and broad calculation accuracy are
unverified.
