## Twenty local PPTX manipulation cases

Bun now edits direct plain-text shapes, appends explicit bullets, sets DrawingML autofit modes and inserts an owned title slide at an exact position. Twenty local cases align staged Python captures with concrete Bun operations; each saves and reopens the package or checks an atomic refusal. Existing reorder, table and notes APIs supply six of the cases.

The selection is 20 cases and 100 steps. [`pptx-manipulation-next20-sources.json`](pptx-manipulation-next20-sources.json) retains every original capture ID, source block, line, file hash and local ID. The source features remain unchanged. These local contracts have no shared execution credit or profile adoption; the shared catalogue remains 304 IDs and 791 cases, with its 732 implemented Bun cases unchanged. The new feature adds 20 local-only cases.

### Production APIs

`Slide.inspectTextShapes()` returns detached shape IDs, names, placeholder kinds, plain text and paragraph levels. Direct shape IDs select edits; no fuzzy matching or grouped-shape traversal occurs.

`Slide.setShapeText(id, text, append?)` resets the selected shape's paragraphs to explicit plain runs. LF separates paragraphs, including empty ones; append adds paragraphs after existing content. It keeps `bodyPr`, `lstStyle`, other shapes and unrelated parts unchanged. This reset deliberately replaces existing direct paragraph/run formatting. User punctuation, bullet characters and whitespace are retained; CR remains a character reference. Unsupported fields, breaks, hyperlink-bearing runs, text locks, duplicate IDs and lexical barriers refuse.

`Slide.addBullet(id, text, level?, boldLabel?)` authors one explicit bullet at level 0..8. A label creates a bold `Label: ` run followed by an unbolded value run. `clearShapeText(id)` leaves one empty paragraph. `setShapeAutofit(id, mode)` writes `normAutofit`, `noAutofit` or `spAutoFit` for `shrink`, `none` or `resize`. It preserves other body properties. No Office rendering, font metrics or overflow quality was measured.

`Presentation.insertTextSlide(index, title, subtitle?)` uses the existing safe title-layout creation and permutation APIs. It validates the original order before adding parts, preserves held slide identities, and restores package bytes, handle order and version snapshots if serialisation fails.

The APIs retain UTF-8 and BOM-marked UTF-16 encoding through the existing package text writer. Shape edits invalidate held paragraph/table anchors only after commit. Refusals leave bytes and versions unchanged.

### Checks

The exact local feature is `features/pptx/manipulation-next20.feature`. Independent XML assertions check text, bullet levels/label flags, autofit elements, table values/grid sums, notes and logical slide order. Caller arrays and unrelated payloads retain custody. Eight native tests cover XML controls, aliases, Unicode/CR/empty paragraphs, UTF-16LE/BE, stale handles and rollback.

Five temporary production faults fail the exact 20-case suite: lost shape edit, wrong bullet level, lost label bold, wrong autofit element and wrong insertion position. Each is restored to 20 cases / 100 steps passing.

```sh
make check
OOXML_FIXTURES_ROOT=/workspace/projects/fixtures-ooxml \
OOXML_REFERENCE_PIN=/workspace/projects/bun-ooxml/docs/behaviors/package-alignment-candidate.json \
make check
```

Default checks pass 1,135 unit tests with 10 candidate-only skips; package-candidate checks pass 1,145 unit tests. Both execute 752/752 acceptance cases and all three Office examples. Reports and failed/restored logs are under `/workspace/analysis/ooxml-pptx-next20-20260930/`.

The published fixture pin remains v0.152.0 / `28e492f50979aaec6ab8d8d001cd9c37790e7fc6`. Shared `14a7bf7` stays untouched for Go's separate package-batch gates. No push, tag, release, pin advancement or central credit occurred.
