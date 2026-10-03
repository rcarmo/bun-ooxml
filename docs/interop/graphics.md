# PowerPoint graphics interoperability

LibreOffice 24.2.7.2 can load, edit, save and reopen the four-slide graphics sample. Microsoft PowerPoint has not been tested. SmartArt interoperability is unverified and the synthetic SmartArt recipes contain schema-invalid graphs.

## Reproduce the LibreOffice server check

The optional development oracle uses Bun to author the sample and Python only for LibreOffice's UNO bindings. Runtime APIs remain Bun-only.

```sh
sudo apt install libreoffice python3-uno
make graphics-uno
```

The graphics tests and sample consume the sealed candidate definitions in the sibling `fixtures-ooxml` checkout. Use shared commit `33a737268bd0b433e1014ac9d287e6db61ec77cd`, or set `OOXML_GRAPHICS_ROOT` to a checkout containing those definitions. The released reference gitlink remains unchanged.

`scripts/graphics-sample.ts` writes `artifacts/graphics-uno/graphics-showcase.pptx`. `scripts/oracles/graphics-uno.py` starts a headless LibreOffice server on a unique local named pipe, with a separate profile and macros disabled. It loads the sample with an interaction handler, edits two labels and moves one shape by 1 mm, exports PPTX and ODP, reopens both and exports PDFs. A 180-second timeout bounds the command. The script terminates its server and writes `artifacts/graphics-uno/uno/report.json`.

## Results on 3 October 2026

Tested the graphics implementation at Bun commit `a81d8128c83d0c67084dd0dd23699effc55c0c3c`, using the shared commit above.

- UNO loaded the sample without an interaction request. Headless loading provides no Microsoft Office repair-dialog evidence.
- Both PPTX and ODP exports reopened with four slides, 23 objects including group children, four pictures, one group and two connectors.
- The edited labels `Input (UNO edited)` and `Grouped A (UNO edited)` survived both exports. Both connectors retained their shape attachments and glue-point indices.
- Picture crop, rotation, picture transparency, shape opacity, gradient colours/stops and inspected line properties survived save/reopen. Import assertions include 20% picture transparency, 35% ellipse transparency and the supplied gradient colours.
- ODP geometry was unchanged. PPTX positions/extents differed by at most one UNO coordinate unit, 0.01 mm. The oracle admits this measured quantisation error; it does not compare generated file hashes. LibreOffice renames title placeholders.

The oracle compares shape order/types, paragraph text, positions/extents, rotation, fill properties, crop, transparency, line properties and connector endpoints. Full preset/path/flip grammars and independent visual fidelity are not covered.

## Independent SDK and conversion checks

Microsoft Open XML SDK 3.5.1 with its Office2019 profile validated 123 of 129 saved outputs captured from 352 passing graphics tests. All outputs in 17 of the 20 families were valid. The practical four-slide sample also validated without errors.

The six failures comprise five synthetic SmartArt outputs and one picture-deletion dependency case. SmartArt inputs contain incomplete layout/style/drawing XML, an invalid extension namespace or unsupported/cycle-bearing relationships. Copying preserves those defects. The deletion case deliberately retains a forbidden presentation-to-image input relationship. No valid Office-produced SmartArt graph was tested.

LibreOffice produced PDFs and PPTX exports for one schema-valid output in each of the 18 non-SmartArt families. PDF page counts matched in all 18 and paragraph text survived all PPTX exports. PDF text markers matched in 17 samples; the diagram contract probe has tiny 1000-by-600-EMU nodes. The practical sample uses legible dimensions and renders the required diagram labels. Generic contract samples changed drawing-object counts on export, so those conversions alone do not establish editable-object preservation.

LibreOffice's PPTX exports add `/docMetadata/LabelInfo.xml` with content type `application/xml`. The SDK expects `application/vnd.ms-office.classificationlabels+xml` and cannot open those exports. The Bun-authored inputs validate independently; UNO can reopen the exported files.

## Microsoft PowerPoint gate

Use a named Microsoft PowerPoint version to open the sample without repair, inspect and edit its graphics, save and reopen it, and check rendering, pictures/SVG fallback, crop, alpha, groups, connectors, freeform geometry and ordering. SmartArt needs a valid Office-produced input and a separate copy/open/save test. Neither SDK validation nor LibreOffice results establish Microsoft application interoperability.
