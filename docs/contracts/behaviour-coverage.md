# Behaviour coverage

Shared reference manifest: **518 assets**. Implementation evidence remains consumer-local.

| Feature | Lifecycle | Scenarios | Expanded cases | Steps |
|---|---|---:|---:|---:|
| features/docx/create.feature | implemented | 4 | 8 | 47 |
| features/docx/redline.feature | implemented | 2 | 6 | 20 |
| features/docx/revisions.feature | implemented | 4 | 4 | 32 |
| features/docx/story.feature | implemented | 3 | 3 | 21 |
| features/docx/tables.feature | implemented | 4 | 7 | 45 |
| features/docx/text.feature | implemented | 5 | 7 | 33 |
| features/opc/graph.feature | implemented | 4 | 4 | 16 |
| features/opc/package.feature | implemented | 3 | 3 | 11 |
| features/opc/relationship-namespaces.feature | implemented | 2 | 4 | 16 |
| features/opc/xml-names.feature | implemented | 3 | 6 | 19 |
| features/opc/xml.feature | implemented | 5 | 5 | 22 |
| features/opc/zip.feature | implemented | 4 | 4 | 25 |
| features/opc/zip64.feature | implemented | 4 | 4 | 13 |
| features/planned/cross-language-followups.feature | planned | 4 | 4 | 17 |
| features/planned/full-port.feature | planned | 5 | 7 | 28 |
| features/planned/office-mutation-additions.feature | planned | 3 | 3 | 12 |
| features/pptx/create.feature | implemented | 3 | 3 | 9 |
| features/pptx/tables.feature | implemented | 4 | 5 | 15 |
| features/pptx/text.feature | implemented | 4 | 4 | 12 |
| features/xlsx/cache-boundaries.feature | implemented | 2 | 3 | 12 |
| features/xlsx/cells.feature | implemented | 8 | 8 | 30 |
| features/xlsx/create.feature | implemented | 7 | 7 | 29 |
| references/fixtures-ooxml/shared/v2/pack/features/mutation-safety.feature | implemented | 8 | 19 | 159 |

Shared workflow IDs and facts are pinned by `references/fixtures-ooxml`.
Counts are contract inventory, not execution evidence. Run `make check` for outcomes.
All planned cases remain gaps. Removing external source inventories does not close behavioural gaps.
