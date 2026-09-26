# Behaviour coverage

The shared reference contains **147 assets**. The table lists available tests; artifacts/acceptance.json records their results.

| Feature | Lifecycle | Scenarios | Expanded cases | Steps |
|---|---|---:|---:|---:|
| features/planned/cross-language-followups.feature | planned | 3 | 3 | 13 |
| features/planned/full-port.feature | planned | 5 | 7 | 28 |
| features/planned/office-mutation-additions.feature | planned | 3 | 3 | 12 |
| references/fixtures-ooxml/workflows/docx/comments.feature | implemented | 4 | 14 | 46 |
| references/fixtures-ooxml/workflows/docx/creation.feature | implemented | 4 | 8 | 47 |
| references/fixtures-ooxml/workflows/docx/effective-formatting.feature | implemented | 2 | 25 | 84 |
| references/fixtures-ooxml/workflows/docx/page-layout.feature | implemented | 2 | 22 | 74 |
| references/fixtures-ooxml/workflows/docx/paragraph-style.feature | implemented | 2 | 19 | 62 |
| references/fixtures-ooxml/workflows/docx/run-formatting.feature | implemented | 2 | 15 | 49 |
| references/fixtures-ooxml/workflows/docx/style-authoring.feature | implemented | 2 | 24 | 86 |
| references/fixtures-ooxml/workflows/docx/tables.feature | implemented | 4 | 7 | 45 |
| references/fixtures-ooxml/workflows/docx/text.feature | implemented | 5 | 7 | 33 |
| references/fixtures-ooxml/workflows/docx/tracked-workflow.feature | implemented | 2 | 17 | 56 |
| references/fixtures-ooxml/workflows/mutation-safety.feature | implemented | 8 | 19 | 159 |
| references/fixtures-ooxml/workflows/native/docx-redline.feature | implemented | 2 | 6 | 20 |
| references/fixtures-ooxml/workflows/native/docx-revisions.feature | implemented | 4 | 4 | 32 |
| references/fixtures-ooxml/workflows/native/docx-story.feature | implemented | 3 | 3 | 21 |
| references/fixtures-ooxml/workflows/native/opc-graph.feature | implemented | 4 | 4 | 16 |
| references/fixtures-ooxml/workflows/native/opc-zip64.feature | implemented | 4 | 4 | 13 |
| references/fixtures-ooxml/workflows/native/pptx-text.feature | implemented | 4 | 4 | 12 |
| references/fixtures-ooxml/workflows/native/xlsx-cache-boundaries.feature | implemented | 2 | 3 | 12 |
| references/fixtures-ooxml/workflows/native/xlsx-cells.feature | implemented | 8 | 8 | 30 |
| references/fixtures-ooxml/workflows/package/preservation.feature | implemented | 3 | 3 | 11 |
| references/fixtures-ooxml/workflows/package/relationship-namespaces.feature | implemented | 2 | 4 | 16 |
| references/fixtures-ooxml/workflows/package/zip32.feature | implemented | 4 | 4 | 25 |
| references/fixtures-ooxml/workflows/pptx/creation.feature | implemented | 3 | 3 | 9 |
| references/fixtures-ooxml/workflows/pptx/slide-order.feature | implemented | 2 | 21 | 70 |
| references/fixtures-ooxml/workflows/pptx/tables.feature | implemented | 4 | 5 | 15 |
| references/fixtures-ooxml/workflows/pptx/text-box.feature | implemented | 2 | 23 | 77 |
| references/fixtures-ooxml/workflows/xlsx/cell-style.feature | implemented | 2 | 27 | 90 |
| references/fixtures-ooxml/workflows/xlsx/creation.feature | implemented | 7 | 7 | 29 |
| references/fixtures-ooxml/workflows/xml/names.feature | implemented | 3 | 6 | 19 |
| references/fixtures-ooxml/workflows/xml/parsing.feature | implemented | 5 | 5 | 22 |

Specifications and shared tests are in `references/fixtures-ooxml`.
Run `make check` to execute implemented scenarios. Planned scenarios are not run.
These counts describe test cases, not the percentage of the OOXML specification supported.
