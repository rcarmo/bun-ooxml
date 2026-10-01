# Direct PPTX formatting candidate

Bun now edits existing shape geometry, direct run and paragraph properties, and slide visibility without rewriting unrelated XML or package members. The candidate executes 20 additional shared cases / 182 steps at `f174097e68e4133ebbfb58daf0f336480d75f210`.

## Production API

- `Slide.hidden` and `setHidden(boolean)` read/write the unqualified slide root `show`. Hide writes 0; show writes 1. Namespaced lookalikes and malformed values refuse. These are OOXML values, with no Office-rendering claim.
- `patchShapeGeometry(shapeId, patch)` patches direct `xfrm` coordinates, extents, rotation and flips. Units are integer EMU and 1/60000-degree angle units. The complete patch is validated before writing; a later invalid extent cannot publish an earlier valid position.
- `patchTextRun(shapeId, paragraph, run, patch)` targets zero-based direct paragraph/run ordinals. Bold/italic false remains an explicit override; null removes the selected direct property. Size is hundredths of a point; font and sRGB edits preserve unpatched properties and language. Missing `rPr` is inserted before text, and children follow schema order.
- `patchParagraph(shapeId, paragraph, patch)` changes alignment, margin/hanging indent, point spacing or percent line spacing. Units are explicit; missing/self-closing `pPr` insertion preserves schema order independently of patch key order.

Edits select a unique ordinary ungrouped shape ID. Unsupported locks, fields, links, extensions, namespace ambiguity, duplicate metadata, malformed colour/spacing or lexical barriers refuse. The existing plain-text replacement policy remains separate. Invalid patches and serialization faults retain session bytes, slide identities and versions.

## Evidence

The 8,634-byte fixture is derived from the previous sealed manipulation input, changing only Body shape paragraphs in `ppt/slides/slide2.xml`. It has two paragraphs and a two-run first paragraph, including rich direct properties and an absent second-run `rPr`. The shared ledger records 17 new project-authored contracts with normative references and three stronger visibility source profiles; no staged alias credit is inferred.

Independent assertions parse saved XML, reopen the package/presentation, check all literal direct values and presence/absence, compare exact changed-member sets and graph identities, and retain every text leaf and unselected property span. A separate consumed-token oracle compares unpatched attribute bytes, including quoted attribute-like strings. Native controls include self-closing/missing-property ordering, Unicode/XML-safe font names, executable patch refusal, locks/ambiguity, late-invalid patches, visibility reorder and serialization rollback.

- Exact selected receipt: 20 cases / 182 steps, with zero failed/undefined/ambiguous/skipped steps.
- Ten native top-level controls.
- Five production faults: wrong visibility, lost geometry/run/paragraph edits, and unrelated attribute drift. Each goes red and passes after restoration.
- Default `make check`: 1,137 passing, 23 skipped, zero failures; 752/752 acceptance.
- Formatting candidate `make check`: 1,155 passing, five prior-lane skips, zero failures; 772/772 acceptance, 59 planned.

```sh
make check
OOXML_FIXTURES_ROOT=/workspace/projects/fixtures-ooxml \
OOXML_REFERENCE_PIN=/workspace/projects/bun-ooxml/docs/behaviors/pptx-formatting-candidate.json \
make check
```

The default reference pin and older XML/package/PPTX candidate files are unchanged. New coverage and outcome reports use separate formatting-candidate paths. No line/branch coverage percentage, PowerPoint rendering, publication, default-pin advancement or central execution credit is claimed. Go and Python need their own reviewed execution receipts.
