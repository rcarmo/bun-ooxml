import {fixturesRoot} from "../../scripts/fixture-inputs.ts";
import {describe,expect,test} from "bun:test";
import {join} from "node:path";
import {OpcPackage} from "../../src/opc/package.ts";
import {Presentation} from "../../src/pptx/index.ts";
import {Workbook} from "../../src/xlsx/index.ts";
import {setCellWrapText} from "../../src/xlsx/styles.ts";
import {OoxmlError} from "../../src/errors.ts";
import {assertPreservedOutput,loadMutationFixtures,type FixtureManifest} from '../../scripts/shared-fixtures.ts';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from "../../scripts/gherkin.ts";
import {namespaceFixture} from "../acceptance/relationship-namespaces.ts";
import {bindings} from '../acceptance/steps.ts';
const pkg='http://schemas.openxmlformats.org/package/2006/relationships';
describe('expanded Office relationship attribute identity',()=>{
 test('PPTX accepts a renamed relationship prefix and preserves it on save',async()=>{
  const bytes=await namespaceFixture('pptx','alias');const deck=await Presentation.open(bytes);
  expect(deck.slides.length).toBe(1);expect(deck.package.text(deck.package.mainPart())).toContain('link:id=');
  expect(deck.package.toBytes()).toEqual(bytes);
 });
 test('XLSX accepts a locally scoped prefix alias through value and style edits',async()=>{
  const bytes=await namespaceFixture('xlsx','alias');const book=await Workbook.open(bytes);
  book.worksheet('Sheet').setCellValue('A1','wrapped\nvalue');setCellWrapText(book,'Sheet','A1',true);
  const reopened=await Workbook.open(book.package.toBytes());expect(reopened.worksheet('Sheet').getCell('A1')?.value).toBe('wrapped\nvalue');expect(reopened.package.text(reopened.package.mainPart())).toContain('link:id=');
  const manifest=await loadMutationFixtures();
  await expect(assertPreservedOutput(book.package.toBytes(),manifest.fixtures.find(f=>f.id==='default-style.xlsx')!)).resolves.toBeUndefined();
 });
 test('wrong URI and unqualified id refuse even with the familiar r spelling',async()=>{
  for(const format of ['pptx','xlsx'] as const)for(const variant of ['wrong','unqualified'] as const){
   const bytes=await namespaceFixture(format,variant);
   await expect(format==='pptx'?Presentation.open(bytes):Workbook.open(bytes)).rejects.toBeInstanceOf(OoxmlError);
  }
 });
 test('worksheet relationship type must use the exact officeDocument URI, not merely its suffix',async()=>{
  const bytes=await namespaceFixture('xlsx','original');const p=await OpcPackage.open(bytes);
  const path='xl/_rels/workbook.xml.rels';p.set(path,p.text(path).replace('http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet','urn:foreign/worksheet'));
  await expect(Workbook.open(p.toBytes())).rejects.toBeInstanceOf(OoxmlError);
 });
 test('locally rebound r takes precedence over its parent declaration',async()=>{
  const original=await namespaceFixture('pptx','original');const package_=await OpcPackage.open(original);const part=package_.mainPart();
  package_.set(part,package_.text(part).replace('<p:sldId ',`<p:sldId xmlns:r="${pkg}" `));
  await expect(Presentation.open(package_.toBytes())).rejects.toThrow('does not resolve to an internal slide relationship');
 });
 test('foreign lookalike id cannot override the real relationship alias',async()=>{
  for(const format of ['pptx','xlsx'] as const){
   const bytes=await namespaceFixture(format,'alias');const p=await OpcPackage.open(bytes);const main=p.mainPart();
   p.set(main,p.text(main).replace(/link:id="([^"]+)"/,`xmlns:r="${pkg}" r:id="bogus" link:id="$1"`));
   if(format==='pptx')expect((await Presentation.open(p.toBytes())).slides.length).toBe(1);
   else expect((await Workbook.open(p.toBytes())).sheetnames).toEqual(['Sheet']);
  }
 });
 test('shared namespace Gherkin executes all four cases with real outcomes',async()=>{
  const path='references/fixtures-ooxml/workflows/package/relationship-namespaces.feature';const feature=parseFeature(path,(await Bun.file(join(fixturesRoot(), "workflows/package/relationship-namespaces.feature")).text()).replace(/^@planned/m, '@implemented @bun'));
  const inventory:AcceptanceInventory={root:'.',features:[feature],counts:{features:{implemented:1,planned:0,total:1},scenarios:{implemented:2,planned:0,total:2},cases:{implemented:4,planned:0,total:4},steps:{implemented:16,planned:0,total:16}}};
  const result=await executeAcceptance(inventory,bindings,newAcceptanceRunId());expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(4);
 });
});
