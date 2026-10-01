# Shared PPTX manipulation candidate

Bun executes the same 20 source cases against shared commit `5dc02ae970e2eaf5ed3e261324ae3961e08bb1c9`, using one immutable three-slide input and 179 concrete steps. The shared ledger preserves the staged source blocks, original IDs, source revisions and Bun local predecessor identities. Go and Python execution needs separate receipts.

The candidate pin is [`pptx-manipulation-candidate.json`](pptx-manipulation-candidate.json): schema v2, explicit root, manifest SHA-256 `70e44ae170ee39175f943bbd1c1a07d64d726bbdcb53507f7be95eec7b21b900` and all 64 feature seals. Shared inventory contains 324 scenarios / 811 cases. The fixture has 20 member payloads and archive SHA-256 `5ad4b68acc5a926c020dbc94154c95fa6c40c7353a55b2f82504879b392e9a45`. Hashes establish input provenance; saved-output acceptance uses concrete values, graph identities, lexical spans and custody comparisons.

## Execution

`tests/acceptance/pptx-manipulation.ts` resolves the fixture by manifest ID. Production shape-text/bullet/autofit/insertion/reorder/table/notes APIs perform the edits. Saved outputs are independently parsed as XML and reopened as packages and presentations. Assertions check exact paragraphs and whitespace, bullet level/character/run bold, the single autofit child, title/subtitle order, native table cells and EMU grid/row sums, notes, exact changed-original-member lists, added slide/layout relationships and retained old identities. Outside the selected XML body/list/insertion, lexical bytes remain unchanged. Invalid permutations retain the session and held slide handles and map `PPTX_ORDER_UNSUPPORTED` to the contract's `invalid-permutation` reason.

The candidate swaps the 20 local predecessor cases for their 20 canonical successors. It retains 752 implemented acceptance cases, 59 planned and no local-only cases. The published default pin still executes 732 shared plus 20 local cases. Older assertion mappings and isolated recipe tests reconstruct only the exact reviewed predecessor feature text. No execution binding uses that historical helper.

## Verified checks

- Exact selected receipt: 20 cases / 179 steps; zero failed, undefined, ambiguous or skipped steps.
- Eight native production controls: Unicode/CR/edge whitespace/empty paragraphs, invalid shape requests, UTF-16, stale anchors and insertion rollback.
- Five temporary production faults: shape no-op, wrong bullet level, wrong autofit, wrong insertion index, unchanged slide order. Each fails the selected suite and passes after restoration.
- Default `make check`: 1,136 passing tests, 12 candidate skips, zero failures; 752/752 acceptance cases.
- PPTX candidate `make check`: 1,145 passing tests, three prior-package-only skips, zero failures; 752/752 acceptance cases.
- Both gates run saved/reopened DOCX/PPTX/XLSX editing/creation and Word redline examples.

```sh
make check
OOXML_FIXTURES_ROOT=/workspace/projects/fixtures-ooxml \
OOXML_REFERENCE_PIN=/workspace/projects/bun-ooxml/docs/behaviors/pptx-manipulation-candidate.json \
make check
```

Evidence is under `/workspace/analysis/ooxml-pptx-crossruntime20-20261001/`. The published default pin remains `28e492f50979aaec6ab8d8d001cd9c37790e7fc6`. Older XML and package candidate files are unchanged. No consumer commit, push, tag, release, default-pin advancement, Office rendering test or central execution credit has occurred in this candidate checkpoint.
