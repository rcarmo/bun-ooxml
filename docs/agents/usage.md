# Using bun-ooxml from an agent

Use the exported classes in `src/index.ts`. Runtime operations run inside Bun.
There is no Python executable, LibreOffice installation or network dependency.

## Decide before editing

1. Keep an immutable input and choose a different output path.
2. Inspect targets and require the intended match count. A missing target is not success.
3. Use a guarded format API. Direct OPC/XML writes require ownership of every affected reference.
4. Save, reopen with the same format reader, and assert the requested value.
5. Compare changed package parts and verify unrelated payloads. XML checks do not prove rendered appearance.

Treat `OoxmlError.code` as a refusal. Do not catch it and continue with raw XML as
an automatic fallback. Return the code and target to the caller. Reacquire stale
spans/anchors after edits rather than retrying them blindly.

## Native batches and previews

```ts
import { patchOffice } from "bun-ooxml";
const changes = [{ target: "<Present>", value: "Approved" }];
const preview = await patchOffice({ source: input, mode: "dry_run", changes });
if (preview.status !== "preview") throw new Error(preview.error?.message);
const result = await patchOffice({
  source: input, output, mode: "safe", changes,
  expectedSourceSha256: preview.sourceSha256,
});
if (result.status !== "committed") throw new Error(result.error?.message);
```

`dry_run` stages in memory but leaves source/destination and directory entries
untouched. `strict` can write the source or a chosen output; `safe` requires a
distinct path. All modes currently require every target to be supported, unique
and non-overlapping. No best-effort fallback occurs. Receipts distinguish matches,
unchanged edits and actually committed changes; refusals always commit zero.

Targets are exact DOCX body/table text, `slide:1/title` or `slide:1/subtitle`, and
`Sheet!A1` or `A1`. Existing cells only; null clears and formula overwrites refuse.
XLSX `multilineWrap: true` clones/reuses the correct xf and preserves its dependencies.
`calculationPolicy: 'invalidate-without-recalculation'` clears cached values on
worksheet formula-bearing cells and reports recalculation required without producing
an answer. Array/data-table result ranges refuse value edits before any write. `expectedDestinationSha256`
can guard an existing output; null requires the output to be absent.

Only this API's writers share its process-local path locks. External editors can
race the final fingerprint check and atomic rename. Choose trusted directories;
leaf symlinks and safe-mode same-file/hardlink destinations refuse. See
`docs/contracts/workflow-api.md` for the bounded target and receipt contract.

## Native creation

```ts
const doc = Document.create();
doc.addParagraph("Heading", { bold: true });
await doc.save("new.docx");

const deck = Presentation.create();
deck.addTextSlide("Title", "Subtitle");
await deck.save("new.pptx");

const wb = Workbook.create(); // Sheet1 and default styles
wb.worksheet("Sheet1").setCellValue("A1", "Value");
wb.addWorksheet("Results").setCellValue("B3", 42);
await wb.save("new.xlsx");
```

These APIs author minimal packages without copying fixtures or invoking Python.
They are not the full authoring surface. DOCX supports simple bold/italic
runs and existing paragraph style IDs; unknown styles refuse. PPTX authors one
owned title layout/master/theme; appending to an existing deck requires one safe,
compatible layout. XLSX can add valid, case-distinct sheet names and missing cells
within Excel bounds, keeping rows/cells sorted and dimension metadata ordered.
Reopen saved files and assert expected content. Native Office rendering has not
been independently verified. `examples/create-office.ts` is an executable example.

## Rectangular tables

```ts
const table = doc.addTable(2, 2);
table.cell(0, 0).text = "Name";
table.cell(1, 1).text = "42";

const slide = deck.slides[0]!;
slide.addTable(2, 2, { x: 0, y: 914400, width: 5486400, height: 1828800 });
slide.tables[0]!.cell(0, 0).text = "Name";
slide.tables[0]!.cell(1, 1).text = "42";
```

Indices are zero-based. PPTX geometry is integer EMU and row heights/column widths
sum to the requested size. Authoring is bounded to 10,000 cells. DOCX table handles
stay usable across cell edits but become stale after structural document edits;
cell handles always become stale after mutation. PPTX table/cell handles become
stale on slide changes: reacquire them from `slide.tables` before the next edit.

DOCX cell replacement retains cell properties and first paragraph/run formatting;
newlines become paragraphs (CRLF/CR normalise to LF). PPTX edits require one plain
supported paragraph and retain first-run formatting. Arbitrary multi-paragraph,
nested table, merged grid, field/drawing or other unsafe topology refuses. There
is no row/column insertion, merging or splitting in this slice. Reopen after save
and assert cell values; visual Office fidelity is not independently verified.

## DOCX

```ts
const doc = await Document.open(input);
const spans = doc.find("Payment is due");
if (spans.length !== 1) throw new Error("Ambiguous or missing target");
await spans[0]!.replace("Payment becomes due");
const outputBytes = await doc.save(); // No implicit overwrite of the input.
const reopened = await Document.open(outputBytes);
if (reopened.find("Payment becomes due").length !== 1) throw new Error("Edit lost");
await doc.save(output);
```

Current search covers body and table-cell paragraphs with supported direct `w:r/w:t` text.
Fields, revisions, controls, hyperlinks, tabs, breaks and other complex topology
may refuse. Run formatting outside the replacement survives. Search is exact and
non-overlapping. `docx-stale-span` means the captured paragraph content changed;
`docx-unsupported-topology` means this slice cannot represent the edit safely.

`Document.find()` can find supported paragraphs while other paragraphs refuse
inspection; absence of matches is not proof that all document stories were searched.
Use `inspectStories()` for linked body/header/footer/note/comment text and examine
its `blindRegions` before relying on coverage. Story-aware search/handles are not
part of `Document.find()`.

## Word review on package snapshots

```ts
import { OpcPackage, inspectStories, trackedReplace, resolveRevisions } from "bun-ooxml";
const pkg = await OpcPackage.open(input);
const inspection = inspectStories(pkg, { view: "current" });
if (inspection.blindRegions.length) throw new Error("Review blind regions first");
trackedReplace(pkg, pkg.mainPart(), "thirty days", "sixty days", {
  author: "Reviewer", date: "2026-09-26T12:00:00Z",
});
await pkg.save(output); // retains native Word insertion/deletion revisions
const accepted = await OpcPackage.open(pkg.toBytes());
resolveRevisions(accepted, "accept", { parts: [accepted.mainPart()] });
```

Review calls operate on `OpcPackage`, not live `Document` caches. Save and reopen
any format readers afterward. `inspectRevisions` reports supported revisions and
unsupported findings without mutation. Resolution preflights every selected story
and refuses unsupported move/format/table/nested markup, protection and unsafe
namespace lifting atomically. It does not filter by author or individual ID.

`trackedReplace` requires exactly one occurrence, including overlapping matches,
within plain direct runs. It rejects unsupported story topology, existing target
revisions, fields, controls, drawings and inter-run barriers. Supply a nonblank
author and valid UTC date. Inserted text uses the first matched run's formatting;
deleted fragments retain their original run properties. Accept/reject text is
verified on private copies before committing. General Word Compare and comment
editing are not implemented. See `docs/contracts/review-api.md` and the executable
`examples/review-word.ts`.

## PPTX

```ts
const deck = await Presentation.open(input);
const slide = deck.slides[0];
if (!slide) throw new Error("No slides");
const matches = slide.inspectText("FY25");
if (matches.length !== 1) throw new Error("Ambiguous or missing paragraph");
slide.replaceTextAt(matches[0]!.anchor, "FY25", "FY26");
await deck.save(output);
const reopened = await Presentation.open(output);
if (reopened.slides[0]!.inspectText("FY26").length !== 1) throw new Error("Edit lost");
```

Slide order follows `sldIdLst` relationships. Anchors belong to the originating
slide instance and version. A replacement changes the first exact match in the
anchored paragraph. `readNotesText()` reads existing notes without creating a part.
Cloning/imports, theme resolution and chart mutation are planned.

## XLSX

```ts
const wb = await Workbook.open(input);
const sheet = wb.worksheet(wb.sheetnames[0]!);
const before = sheet.getCell("A2");
sheet.setCellValue("A2", "Updated value");
await wb.save(output); // Always provide a new path: omitted path uses the opened path.
const reopened = await Workbook.open(output);
if (reopened.worksheet(wb.sheetnames[0]!).getCell("A2")?.value !== "Updated value") {
  throw new Error("Cell edit lost");
}
```

Supported simple cells can be written or added to ordinary worksheets. Unsafe
row/cell order and unsupported structures refuse insertion. Numeric, boolean, inline and
shared-string values can be read; formulas expose stored cached values. The API
preserves the target cell style index. Shared and array formula overwrites refuse.
After an input edit, this slice clears `<v>` contents only on worksheet cells
carrying `<f>`, across all loaded worksheets, and requests recalculation on open.
Array/data-table formulas may have result followers without `<f>`; value edits in
such workbooks refuse atomically with `xlsx-cache-topology-unsupported`. Style-only
edits do not need that invalidation and preserve caches.

Chart caches, external-link caches, other opaque derived values and calculation
chains are preserved, not refreshed or certified. A recalculation flag or receipt
does not establish their freshness. This API does not calculate formulas. A missing
cached value is not zero, a current answer or a successful calculation.

## Guarded package graph edits

```ts
import { addPart, addRelationship, diffPackages } from "bun-ooxml/opc";
const before = await OpcPackage.open(input);
const edited = await OpcPackage.open(before.toBytes());
addPart(edited, "custom/data.bin", new Uint8Array([1, 2, 3]), "application/octet-stream");
addRelationship(edited, "", "urn:example:data", "custom/data.bin");
const report = diffPackages(before, edited);
await edited.save(output);
```

Graph helpers mutate a valid package within a synchronous rollback boundary and
validate before returning. They do not update format object-model caches; use them
before opening a format reader or reopen afterward. Never modify an already-loaded
format's graph behind cached handles and assume its model automatically refreshes.

`removePart` refuses incoming relationships. Explicitly detach an unreferenced
relationship first; removal also refuses if the owner's XML uses its officeDocument
`id`, `embed` or `link`. Unknown binary owner link semantics refuse. No cascade or
unrelated orphan collection occurs. `nextPartName` checks all preserved members,
including case-colliding orphan names. `walkParts` follows internal edges once and
never fetches external targets.

`diffPackages` reports SHA256, byte length and effective content types, including
type-only changes. It ignores ZIP metadata differences and does not claim XML
semantic equivalence, rendered appearance or calculation correctness. See
`docs/contracts/graph-api.md` for ownership limits and the shared workflow references.

## Package-level operations

`OpcPackage` retains original bytes, returns detached part copies, resolves internal
relationships without network access, and rejects dangling references on save.
`diff()` is cumulative relative to the opened package. `transaction()` is for
synchronous work only. A callback must not schedule timers, promises or detached work.
Path saves validate first, fsync a sibling temporary file and atomically rename it.
A symlink at the final destination refuses. Parent directories can be symlinks;
callers must choose a trusted output directory. This API is not a filesystem sandbox
and cannot defend against a concurrently hostile filesystem owner. No-op output
retains the exact original archive bytes.

ZIP defaults: 256 MiB archive, 128 MiB per entry, 512 MiB expanded total, 10,000
entries, maximum compression ratio 1,000. Pass stricter `ZipLimits` to package
loading for untrusted inputs. XML has separate fixed input/depth/node limits.
Single-disk ZIP64 is supported with the same bounds, validated 64-bit metadata and
hard inflate caps. Multi-disk archives, ambiguous extras and extensible ZIP64 end
sectors refuse. The writer is in-memory, not a large-file streaming API. Unsupported
formats must never be converted silently.

See `examples/agent-edit.ts` for fixture-backed editing and `examples/create-office.ts`
for authoring without fixture inputs.
