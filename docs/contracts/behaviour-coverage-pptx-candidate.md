# Behaviour coverage

The shared reference contains **351 assets**. The table lists available tests; artifacts/acceptance.json records their results.

| Scope | Scenarios | Cases | Implemented cases | Planned cases |
|---|---:|---:|---:|---:|
| Shared catalogue | 324 | 811 | 752 | 59 |
| Local-only obligations | 0 | 0 | 0 | 0 |
| Combined inventory | 324 | 811 | 752 | 59 |

The canonical denominator is **324 shared scenarios / 811 shared cases**. Local-only obligations are additional and are not included in that denominator. Unselected canonical features remain planned; inventory presence is not execution evidence.

| Feature | Lifecycle | Scenarios | Expanded cases | Implemented cases | Planned cases | Steps |
|---|---|---:|---:|---:|---:|---:|
| references/fixtures-ooxml/workflows/docx/anchor-discovery.feature | mixed | 5 | 5 | 4 | 1 | 23 |
| references/fixtures-ooxml/workflows/docx/comment-content-type.feature | planned | 1 | 1 | 0 | 1 | 4 |
| references/fixtures-ooxml/workflows/docx/comments.feature | mixed | 21 | 46 | 37 | 9 | 210 |
| references/fixtures-ooxml/workflows/docx/completion-audit.feature | planned | 3 | 4 | 0 | 4 | 36 |
| references/fixtures-ooxml/workflows/docx/creation.feature | implemented | 5 | 9 | 9 | 0 | 50 |
| references/fixtures-ooxml/workflows/docx/effective-formatting.feature | implemented | 2 | 25 | 25 | 0 | 84 |
| references/fixtures-ooxml/workflows/docx/font-size.feature | implemented | 1 | 1 | 1 | 0 | 6 |
| references/fixtures-ooxml/workflows/docx/mutation-safety.feature | implemented | 3 | 7 | 7 | 0 | 52 |
| references/fixtures-ooxml/workflows/docx/page-layout.feature | implemented | 3 | 23 | 23 | 0 | 77 |
| references/fixtures-ooxml/workflows/docx/paragraph-style.feature | mixed | 3 | 24 | 19 | 5 | 82 |
| references/fixtures-ooxml/workflows/docx/paragraphs.feature | implemented | 6 | 16 | 16 | 0 | 49 |
| references/fixtures-ooxml/workflows/docx/properties.feature | implemented | 1 | 1 | 1 | 0 | 7 |
| references/fixtures-ooxml/workflows/docx/review-integration.feature | planned | 2 | 2 | 0 | 2 | 9 |
| references/fixtures-ooxml/workflows/docx/revisions.feature | implemented | 23 | 88 | 88 | 0 | 396 |
| references/fixtures-ooxml/workflows/docx/run-formatting.feature | mixed | 10 | 43 | 42 | 1 | 137 |
| references/fixtures-ooxml/workflows/docx/stories.feature | implemented | 3 | 3 | 3 | 0 | 21 |
| references/fixtures-ooxml/workflows/docx/style-authoring.feature | implemented | 2 | 24 | 24 | 0 | 86 |
| references/fixtures-ooxml/workflows/docx/table-merging.feature | implemented | 14 | 52 | 52 | 0 | 193 |
| references/fixtures-ooxml/workflows/docx/tables.feature | mixed | 14 | 24 | 23 | 1 | 104 |
| references/fixtures-ooxml/workflows/docx/template-analysis.feature | mixed | 14 | 28 | 22 | 6 | 112 |
| references/fixtures-ooxml/workflows/docx/template-cache.feature | planned | 7 | 7 | 0 | 7 | 38 |
| references/fixtures-ooxml/workflows/docx/text.feature | implemented | 5 | 7 | 7 | 0 | 33 |
| references/fixtures-ooxml/workflows/docx/tracked-workflow.feature | implemented | 3 | 18 | 18 | 0 | 60 |
| references/fixtures-ooxml/workflows/docx/tracking-settings.feature | implemented | 7 | 24 | 24 | 0 | 101 |
| references/fixtures-ooxml/workflows/office/full-coverage.feature | planned | 1 | 3 | 0 | 3 | 12 |
| references/fixtures-ooxml/workflows/package/admission-limit-configuration.feature | implemented | 1 | 2 | 2 | 0 | 12 |
| references/fixtures-ooxml/workflows/package/data-descriptor-integrity.feature | implemented | 1 | 1 | 1 | 0 | 8 |
| references/fixtures-ooxml/workflows/package/graph.feature | implemented | 4 | 4 | 4 | 0 | 17 |
| references/fixtures-ooxml/workflows/package/preservation.feature | implemented | 10 | 14 | 14 | 0 | 120 |
| references/fixtures-ooxml/workflows/package/relationship-namespaces.feature | implemented | 2 | 4 | 4 | 0 | 16 |
| references/fixtures-ooxml/workflows/package/semantic-diff.feature | implemented | 1 | 1 | 1 | 0 | 7 |
| references/fixtures-ooxml/workflows/package/xml-member-admission.feature | implemented | 1 | 3 | 3 | 0 | 9 |
| references/fixtures-ooxml/workflows/package/zip-admission.feature | implemented | 4 | 11 | 11 | 0 | 38 |
| references/fixtures-ooxml/workflows/package/zip32.feature | implemented | 8 | 24 | 24 | 0 | 118 |
| references/fixtures-ooxml/workflows/package/zip64.feature | implemented | 4 | 4 | 4 | 0 | 13 |
| references/fixtures-ooxml/workflows/pptx/bullets.feature | implemented | 5 | 5 | 5 | 0 | 45 |
| references/fixtures-ooxml/workflows/pptx/creation.feature | implemented | 5 | 5 | 5 | 0 | 27 |
| references/fixtures-ooxml/workflows/pptx/layout-recommendation.feature | planned | 2 | 6 | 0 | 6 | 39 |
| references/fixtures-ooxml/workflows/pptx/mutation-safety.feature | implemented | 3 | 6 | 6 | 0 | 45 |
| references/fixtures-ooxml/workflows/pptx/notes.feature | mixed | 12 | 13 | 11 | 2 | 87 |
| references/fixtures-ooxml/workflows/pptx/preservation.feature | implemented | 1 | 1 | 1 | 0 | 6 |
| references/fixtures-ooxml/workflows/pptx/slide-import.feature | planned | 1 | 1 | 0 | 1 | 4 |
| references/fixtures-ooxml/workflows/pptx/slide-order.feature | implemented | 4 | 23 | 23 | 0 | 87 |
| references/fixtures-ooxml/workflows/pptx/slide-visibility.feature | planned | 2 | 2 | 0 | 2 | 16 |
| references/fixtures-ooxml/workflows/pptx/tables.feature | implemented | 6 | 7 | 7 | 0 | 33 |
| references/fixtures-ooxml/workflows/pptx/text-autofit.feature | implemented | 3 | 3 | 3 | 0 | 27 |
| references/fixtures-ooxml/workflows/pptx/text-box.feature | implemented | 2 | 23 | 23 | 0 | 77 |
| references/fixtures-ooxml/workflows/pptx/text.feature | implemented | 7 | 7 | 7 | 0 | 45 |
| references/fixtures-ooxml/workflows/xlsx/cache-completeness.feature | planned | 1 | 1 | 0 | 1 | 4 |
| references/fixtures-ooxml/workflows/xlsx/calculation-chain-lifecycle.feature | implemented | 1 | 1 | 1 | 0 | 14 |
| references/fixtures-ooxml/workflows/xlsx/calculation-engine.feature | planned | 1 | 1 | 0 | 1 | 4 |
| references/fixtures-ooxml/workflows/xlsx/cell-style.feature | implemented | 3 | 28 | 28 | 0 | 101 |
| references/fixtures-ooxml/workflows/xlsx/cells.feature | implemented | 7 | 7 | 7 | 0 | 26 |
| references/fixtures-ooxml/workflows/xlsx/comment-vml-custody.feature | mixed | 6 | 6 | 1 | 5 | 29 |
| references/fixtures-ooxml/workflows/xlsx/creation.feature | implemented | 7 | 7 | 7 | 0 | 29 |
| references/fixtures-ooxml/workflows/xlsx/formula-cache.feature | implemented | 4 | 5 | 5 | 0 | 30 |
| references/fixtures-ooxml/workflows/xlsx/formula-references.feature | implemented | 9 | 45 | 45 | 0 | 152 |
| references/fixtures-ooxml/workflows/xlsx/mutation-safety.feature | implemented | 2 | 6 | 6 | 0 | 45 |
| references/fixtures-ooxml/workflows/xlsx/structural-edits.feature | planned | 1 | 1 | 0 | 1 | 4 |
| references/fixtures-ooxml/workflows/xlsx/style-readback.feature | implemented | 1 | 1 | 1 | 0 | 4 |
| references/fixtures-ooxml/workflows/xml/comparison.feature | implemented | 5 | 10 | 10 | 0 | 40 |
| references/fixtures-ooxml/workflows/xml/editing.feature | implemented | 11 | 16 | 16 | 0 | 61 |
| references/fixtures-ooxml/workflows/xml/names.feature | implemented | 3 | 6 | 6 | 0 | 20 |
| references/fixtures-ooxml/workflows/xml/parsing.feature | implemented | 14 | 15 | 15 | 0 | 64 |

Specifications and shared tests are in `references/fixtures-ooxml`.
Run `make check` to execute implemented scenarios. Planned scenarios are not run.
These counts describe test cases, not the percentage of the OOXML specification supported.
