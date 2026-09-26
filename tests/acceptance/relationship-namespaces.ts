import {expect} from 'bun:test';
import {join} from 'node:path';
import {OpcPackage} from '../../src/opc/package.ts';
import {Workbook} from '../../src/xlsx/index.ts';
import {setCellWrapText} from '../../src/xlsx/styles.ts';
import {Presentation} from '../../src/pptx/index.ts';
import {OoxmlError} from '../../src/errors.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
const office='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const packageUri='http://schemas.openxmlformats.org/package/2006/relationships';
export async function namespaceFixture(format:'pptx'|'xlsx',variant:'original'|'alias'|'wrong'|'unqualified'):Promise<Uint8Array>{
 const name=format==='pptx'?'title-and-subtitle.pptx':'default-style.xlsx';
 const p=await OpcPackage.open(join(import.meta.dir,'../../docs/contracts/shared-v2/pack/fixtures',name));const part=p.mainPart();let xml=p.text(part);
 if(variant==='alias')xml=xml.replaceAll('xmlns:r=','xmlns:link=').replaceAll('r:id=','link:id=');
 if(variant==='wrong')xml=xml.replaceAll(office,packageUri);
 if(variant==='unqualified')xml=xml.replaceAll('r:id=','id=');
 p.set(part,xml);return p.toBytes();
}
export const bindings:StepBinding[]=[
 {pattern:/^the shared (pptx|xlsx) fixture with its relationship prefix changed to link$/,run:async(ctx,format)=>{ctx.format=format;ctx.input=await namespaceFixture(format as 'pptx'|'xlsx','alias');}},
 {pattern:/^Bun opens it edits its text and saves then reopens it$/,run:async ctx=>{
  if(ctx.format==='pptx'){const p=await Presentation.open(ctx.input as Uint8Array);const t=p.slides[0]!.inspectText('')[0]!;p.slides[0]!.replaceTextAt(t.anchor,t.text,'Edited title');ctx.output=p.package.toBytes();ctx.reopened=await Presentation.open(ctx.output as Uint8Array);}
  else {const w=await Workbook.open(ctx.input as Uint8Array);w.worksheet('Sheet').setCellValue('A1','Edited\ncell');setCellWrapText(w,'Sheet','A1',true);ctx.output=w.package.toBytes();ctx.reopened=await Workbook.open(ctx.output as Uint8Array);}
 }},
 {pattern:/^the requested (pptx|xlsx) value survives and all relationship targets resolve$/,run:async(ctx,format)=>{
  if(format==='pptx')expect((ctx.reopened as Presentation).slides[0]!.inspectText('')[0]!.text).toBe('Edited title');
  else expect((ctx.reopened as Workbook).worksheet('Sheet').getCell('A1')?.value).toBe('Edited\ncell');
  await expect(OpcPackage.open(ctx.output as Uint8Array)).resolves.toBeInstanceOf(OpcPackage);
 }},
 {pattern:/^the original relationship attribute spelling is preserved$/,run:async ctx=>{const p=await OpcPackage.open(ctx.output as Uint8Array);expect(p.text(p.mainPart())).toContain('link:id=');expect(p.text(p.mainPart())).not.toContain(' r:id=');}},
 {pattern:/^the shared (pptx|xlsx) fixture with r bound to package relationships$/,run:async(ctx,format)=>{ctx.format=format;ctx.input=await namespaceFixture(format as 'pptx'|'xlsx','wrong');ctx.before=(ctx.input as Uint8Array).slice();}},
 {pattern:/^Bun attempts to open the namespace-mismatched Office document$/,run:async ctx=>{try{if(ctx.format==='pptx')await Presentation.open(ctx.input as Uint8Array);else await Workbook.open(ctx.input as Uint8Array);}catch(e){ctx.error=e;}}},
 {pattern:/^the (pptx|xlsx) reader refuses with its structural error code$/,run:(ctx,format)=>{expect(ctx.error).toBeInstanceOf(OoxmlError);expect((ctx.error as OoxmlError).code).toBe(format==='pptx'?'PPTX_PRESENTATION_INVALID':'xlsx-workbook-invalid');}},
 {pattern:/^the supplied archive bytes remain unchanged$/,run:ctx=>{expect(ctx.input as Uint8Array).toEqual(ctx.before as Uint8Array);}},
];
