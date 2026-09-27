# Behaviour coverage

The shared reference contains **202 assets**. The table lists available tests; artifacts/acceptance.json records their results.

| Feature | Lifecycle | Scenarios | Expanded cases | Implemented cases | Planned cases | Steps |
|---|---|---:|---:|---:|---:|---:|
| features/planned/cross-language-followups.feature | planned | 3 | 3 | 0 | 3 | 13 |
| features/planned/full-port.feature | planned | 5 | 7 | 0 | 7 | 28 |
| features/planned/office-mutation-additions.feature | planned | 1 | 1 | 0 | 1 | 4 |
| references/fixtures-ooxml/workflows/docx/comments.feature | mixed | 13 | 23 | 14 | 9 | 102 |
| references/fixtures-ooxml/workflows/docx/creation.feature | implemented | 4 | 8 | 8 | 0 | 47 |
| references/fixtures-ooxml/workflows/docx/document-model.feature | mixed | 29 | 70 | 61 | 9 | 233 |
| references/fixtures-ooxml/workflows/docx/effective-formatting.feature | implemented | 2 | 25 | 25 | 0 | 84 |
| references/fixtures-ooxml/workflows/docx/font-size.feature | implemented | 1 | 1 | 1 | 0 | 6 |
| references/fixtures-ooxml/workflows/docx/page-layout.feature | implemented | 2 | 22 | 22 | 0 | 74 |
| references/fixtures-ooxml/workflows/docx/paragraph-style.feature | implemented | 2 | 19 | 19 | 0 | 62 |
| references/fixtures-ooxml/workflows/docx/run-formatting.feature | implemented | 2 | 15 | 15 | 0 | 49 |
| references/fixtures-ooxml/workflows/docx/style-authoring.feature | implemented | 2 | 24 | 24 | 0 | 86 |
| references/fixtures-ooxml/workflows/docx/tables.feature | implemented | 4 | 7 | 7 | 0 | 45 |
| references/fixtures-ooxml/workflows/docx/text.feature | implemented | 5 | 7 | 7 | 0 | 33 |
| references/fixtures-ooxml/workflows/docx/tracked-workflow.feature | implemented | 2 | 17 | 17 | 0 | 56 |
| references/fixtures-ooxml/workflows/mutation-safety.feature | implemented | 8 | 19 | 19 | 0 | 159 |
| references/fixtures-ooxml/workflows/native/docx-redline.feature | implemented | 2 | 6 | 6 | 0 | 20 |
| references/fixtures-ooxml/workflows/native/docx-revisions.feature | implemented | 4 | 4 | 4 | 0 | 32 |
| references/fixtures-ooxml/workflows/native/docx-story.feature | implemented | 3 | 3 | 3 | 0 | 21 |
| references/fixtures-ooxml/workflows/native/opc-graph.feature | implemented | 4 | 4 | 4 | 0 | 16 |
| references/fixtures-ooxml/workflows/native/opc-zip64.feature | implemented | 4 | 4 | 4 | 0 | 13 |
| references/fixtures-ooxml/workflows/native/pptx-text.feature | mixed | 13 | 13 | 12 | 1 | 58 |
| references/fixtures-ooxml/workflows/native/xlsx-cache-boundaries.feature | implemented | 2 | 3 | 3 | 0 | 12 |
| references/fixtures-ooxml/workflows/native/xlsx-cells.feature | implemented | 8 | 8 | 8 | 0 | 30 |
| references/fixtures-ooxml/workflows/package/preservation.feature | implemented | 10 | 14 | 14 | 0 | 119 |
| references/fixtures-ooxml/workflows/package/relationship-namespaces.feature | implemented | 2 | 4 | 4 | 0 | 16 |
| references/fixtures-ooxml/workflows/package/semantic-diff.feature | implemented | 1 | 1 | 1 | 0 | 7 |
| references/fixtures-ooxml/workflows/package/xml-member-admission.feature | implemented | 1 | 3 | 3 | 0 | 9 |
| references/fixtures-ooxml/workflows/package/zip-admission.feature | implemented | 3 | 10 | 10 | 0 | 31 |
| references/fixtures-ooxml/workflows/package/zip32.feature | implemented | 8 | 24 | 24 | 0 | 116 |
| references/fixtures-ooxml/workflows/pptx/creation.feature | implemented | 3 | 3 | 3 | 0 | 9 |
| references/fixtures-ooxml/workflows/pptx/slide-order.feature | implemented | 2 | 21 | 21 | 0 | 70 |
| references/fixtures-ooxml/workflows/pptx/tables.feature | implemented | 4 | 5 | 5 | 0 | 15 |
| references/fixtures-ooxml/workflows/pptx/text-box.feature | implemented | 2 | 23 | 23 | 0 | 77 |
| references/fixtures-ooxml/workflows/workflow-receipts.feature | implemented | 2 | 2 | 2 | 0 | 8 |
| references/fixtures-ooxml/workflows/xlsx/cell-style.feature | implemented | 2 | 27 | 27 | 0 | 90 |
| references/fixtures-ooxml/workflows/xlsx/creation.feature | implemented | 7 | 7 | 7 | 0 | 29 |
| references/fixtures-ooxml/workflows/xlsx/formula-references.feature | implemented | 9 | 45 | 45 | 0 | 152 |
| references/fixtures-ooxml/workflows/xml/comparison.feature | implemented | 5 | 10 | 10 | 0 | 40 |
| references/fixtures-ooxml/workflows/xml/names.feature | implemented | 3 | 6 | 6 | 0 | 19 |
| references/fixtures-ooxml/workflows/xml/parsing.feature | implemented | 25 | 31 | 31 | 0 | 116 |

Specifications and shared tests are in `references/fixtures-ooxml`.
Run `make check` to execute implemented scenarios. Planned scenarios are not run.
These counts describe test cases, not the percentage of the OOXML specification supported.
