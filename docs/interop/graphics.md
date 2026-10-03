# PowerPoint graphics interoperability

LibreOffice 24.2.7.2 can load, edit, save and reopen the four-slide graphics sample. Microsoft PowerPoint has not been tested. A sealed PowerPoint-produced SmartArt input and its isolated copies now pass SDK validation and LibreOffice import/save/reopen. The older synthetic SmartArt recipes still contain schema-invalid graphs.

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

The six failures comprise five synthetic SmartArt outputs and one picture-deletion dependency case. SmartArt inputs contain incomplete layout/style/drawing XML, an invalid extension namespace or unsupported/cycle-bearing relationships. Copying preserves those defects. The deletion case deliberately retains a forbidden presentation-to-image input relationship. That capture predates the Office source checks below.

LibreOffice produced PDFs and PPTX exports for one schema-valid output in each of the 18 non-SmartArt families. PDF page counts matched in all 18 and paragraph text survived all PPTX exports. PDF text markers matched in 17 samples; the diagram contract probe has tiny 1000-by-600-EMU nodes. The practical sample uses legible dimensions and renders the required diagram labels. Generic contract samples changed drawing-object counts on export, so those conversions alone do not establish editable-object preservation.

LibreOffice's PPTX exports add `/docMetadata/LabelInfo.xml` with content type `application/xml`. The SDK expects `application/vnd.ms-office.classificationlabels+xml` and cannot open those exports. The Bun-authored inputs validate independently; UNO can reopen the exported files.

## SmartArt Office source checks

The shared candidate fixture `fixture-b97e4c6d2ee1dd4094f50f9043820610268dd452d7717c537f5456636adcc353` contains unchanged Apache POI `test-data/slideshow/SmartArt.pptx` bytes from commit `33f89110dd72c7b94710322ce7e795d0f464f68f`. Its metadata reports Microsoft Office PowerPoint. The shared repository retains its Apache-2.0 licence, notice, Git blob identity, source revision and member hashes.

This input exposed unsupported Office encoding: `dataModelExt` resolves its drawing relationship on the slide, layout templates have numeric sample identities, instance transitions use `cxnId`, and persisted drawing identities can repeat zero. Inspection now follows the slide-owned drawing edge. Copying preserves template identities, remaps instance GUIDs/references, allocates unique drawing IDs and updates the copied metadata relationship atomically.

Three new native cases cover inspection, cross-presentation copy and same-slide copy. All 27 SmartArt tests pass, including atomic refusal and rollback controls. The source and both copies pass Microsoft Open XML SDK 3.5.1 Office2019 validation. LibreOffice UNO loads, saves and reopens all three without an interaction request, preserving one group for source/cross-presentation copy and two groups for same-slide copy, with six children each. Group positions are unchanged; extents differ by at most 0.01 mm. The source drawing contains no text labels, so no SmartArt text-editing result is available.

```sh
make smartart-uno
# Optional independent SDK gate after building the existing schema oracle:
dotnet tests/oracles/schema/bin/Release/net10.0/SchemaCheck.dll \\
  artifacts/smartart-uno/source.pptx \\
  artifacts/smartart-uno/cross-presentation-copy.pptx \\
  artifacts/smartart-uno/same-slide-copy.pptx
```

The candidate shared root must contain `ledgers/pptx-smartart-office-source.json`; the release pin is unchanged. Another exploratory POI rotated-text sample is still refused for a dangling model reference. These results cover one admitted source encoding, not arbitrary Office SmartArt.

## Microsoft PowerPoint gate

Use a named Microsoft PowerPoint version to open the sample without repair, inspect and edit its graphics, save and reopen it, and check rendering, pictures/SVG fallback, crop, alpha, groups, connectors, freeform geometry and ordering. SmartArt has a valid Office-produced input with SDK and LibreOffice evidence; it still needs an actual Microsoft PowerPoint copy/open/save test. Neither SDK validation nor LibreOffice results establish Microsoft application interoperability.
