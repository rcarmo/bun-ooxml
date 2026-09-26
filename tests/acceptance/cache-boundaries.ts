import {expect} from 'bun:test';
import {createCrossSheetCachedFormulaWorkbook} from './xlsx.ts';
import {Workbook} from '../../src/xlsx/index.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';
import {OoxmlError} from '../../src/errors.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
const encode=(s:string)=>new TextEncoder().encode(s);
const decode=(b:Uint8Array)=>new TextDecoder().decode(b);
export function rangedFormulaFixture(topology:'array'|'dataTable'):Uint8Array{
 const parts=readZip(createCrossSheetCachedFormulaWorkbook());
 parts.set('xl/worksheets/sheet2.xml',encode(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><f t="${topology}" ref="A1:A2">Model!A1*{1;2}</f><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row></sheetData></worksheet>`));
 return writeZip(parts);
}
export function opaqueCacheFixture():Uint8Array{
 const parts=readZip(createCrossSheetCachedFormulaWorkbook());
 // Reference-only opaque cache payloads are linked from the package; no chart
 // authoring or external-link update semantics are asserted by this fixture.
 parts.set('xl/charts/cache-boundary.xml',encode('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:numCache><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:chartSpace>'));
 parts.set('xl/externalLinks/cache-boundary.xml',encode('<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook><sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>1</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>'));
 parts.set('_rels/.rels',encode(decode(parts.get('_rels/.rels')!).replace('</Relationships>','<Relationship Id="cache1" Type="urn:cache-fixture/chart" Target="xl/charts/cache-boundary.xml"/><Relationship Id="cache2" Type="urn:cache-fixture/external" Target="xl/externalLinks/cache-boundary.xml"/></Relationships>')));
 return writeZip(parts);
}
export const bindings:StepBinding[]=[
 {pattern:/^a two-sheet workbook with a (array|dataTable) formula and cached follower cells$/,run:async(ctx,topology)=>{ctx.input=rangedFormulaFixture(topology as 'array'|'dataTable');ctx.book=await Workbook.open(ctx.input as Uint8Array);}},
 {pattern:/^an ordinary input value edit is attempted$/,run:ctx=>{try{(ctx.book as Workbook).worksheet('Model').setCellValue('A1',10);}catch(e){ctx.refusal=e;}}},
 {pattern:/^the edit refuses with code "(xlsx-cache-topology-unsupported)"$/,run:(ctx,code)=>{expect(ctx.refusal).toBeInstanceOf(OoxmlError);expect((ctx.refusal as OoxmlError).code).toBe(code);}},
 {pattern:/^every workbook member and cached result remains byte-identical$/,run:ctx=>{expect((ctx.book as Workbook).toBytes()).toEqual(ctx.input as Uint8Array);}},
 {pattern:/^a workbook with ordinary worksheet formulas and opaque chart and external-link caches$/,run:async ctx=>{ctx.input=opaqueCacheFixture();ctx.book=await Workbook.open(ctx.input as Uint8Array);}},
 {pattern:/^its input is changed and the workbook is saved and reopened$/,run:async ctx=>{const w=ctx.book as Workbook;w.worksheet('Model').setCellValue('A1',10);ctx.output=w.toBytes();ctx.reopened=await Workbook.open(ctx.output as Uint8Array);}},
 {pattern:/^worksheet formula cache values are cleared and their formulas retained$/,run:ctx=>{const w=ctx.reopened as Workbook;expect(w.worksheet('Model').getCell('B1')).toMatchObject({kind:'formula',formula:'A1+A2',cached:null});expect(w.worksheet('Summary').getCell('A1')).toMatchObject({kind:'formula',formula:'Model!B1*3',cached:null});}},
 {pattern:/^the opaque chart and external-link cache payloads remain unchanged$/,run:ctx=>{const before=readZip(ctx.input as Uint8Array),after=readZip(ctx.output as Uint8Array);for(const path of ['xl/charts/cache-boundary.xml','xl/externalLinks/cache-boundary.xml'])expect(after.get(path)).toEqual(before.get(path)!);}},
];
