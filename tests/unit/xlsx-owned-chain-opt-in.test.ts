import {expect,test} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {fixturePath,F} from '../../scripts/fixture-inputs.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';
import {OpcPackage} from '../../src/opc/package.ts';
import {Workbook} from '../../src/xlsx/index.ts';
import {patchOffice} from '../../src/workflow/index.ts';
import {attribute,elements,parseXml} from '../../src/xml/index.ts';

const chain='xl/chains/order.xml',rel='xl/_rels/workbook.xml.rels',types='[Content_Types].xml';
const relationship='http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
const contentType='application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
const encode=new TextEncoder(),decode=new TextDecoder();
const allowed=new Set(['xl/workbook.xml','xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml',rel,types]);
const spreadsheetNamespace='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
// Read stored values from the saved archive without invoking the workbook reader
// or calculating any formula. A missing or empty <v> has no data-only answer.
function savedFormula(bytes:Uint8Array,part:string,reference:string){
 const xml=parseXml(decode.decode(readZip(bytes).get(part)!));
 const cell=elements(xml.root,'c',spreadsheetNamespace).find(c=>attribute(c,'r')===reference);
 expect(cell).toBeDefined();
 const formula=cell!.children.find(e=>e.localName==='f'&&e.namespaceURI===spreadsheetNamespace);
 const value=cell!.children.find(e=>e.localName==='v'&&e.namespaceURI===spreadsheetNamespace);
 return {formula:formula?.text, dataOnly:value?.text||null};
}
export async function ownedChainSource(){
 const sealed=Uint8Array.from(await Bun.file(fixturePath(F.shared.crossSheetCache)).bytes()),parts=readZip(sealed);
 const sheet='xl/worksheets/sheet2.xml',xml=decode.decode(parts.get(sheet)!);
 expect(xml).toContain('<s:f>Input!A1*2</s:f><s:v>2</s:v>');
 parts.set(sheet,encode.encode(xml.replace('</s:row>','<s:c r="B1"><s:f>A1+1</s:f><s:v>3</s:v></s:c><s:c r="C1"><s:f>42</s:f><s:v>42</s:v></s:c></s:row>')));
 const relXml=decode.decode(parts.get(rel)!);expect(relXml).not.toContain(relationship);
 parts.set(rel,encode.encode(relXml.replace('</rel:Relationships>',`<rel:Relationship Id="rIdOwnedCalcChain" Type="${relationship}" Target="chains/order.xml"/></rel:Relationships>`)));
 const typesXml=decode.decode(parts.get(types)!);
 parts.set(types,encode.encode(typesXml.replace('</ct:Types>',`<ct:Override PartName="/${chain}" ContentType="${contentType}"/></ct:Types>`)));
 parts.set(chain,encode.encode('<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="2"/><c r="B1" i="2"/></calcChain>'));
 return writeZip(parts);
}

test('opt-in dependent invalidation removes one owned chain and preserves independent cache',async()=>{
 const source=await ownedChainSource(),root=await mkdtemp(join(tmpdir(),'bun-chain-positive-'));
 try{
  const input=join(root,'input.xlsx'),output=join(root,'out.xlsx');await Bun.write(input,source);
  const result=await patchOffice({source:input,output,mode:'safe',calculationPolicy:'invalidate-dependent-formula-caches',changes:[{target:'Input!A1',value:10}]});
  expect(result.error).toBeUndefined();expect(result.status).toBe('committed');expect(result.committedChanges).toBe(1);expect(result.calculationState).toBe('recalculation-required');
  expect([...await Bun.file(input).bytes()]).toEqual([...source]);
  const bytes=Uint8Array.from(await Bun.file(output).bytes()),reopened=await Workbook.open(bytes),before=readZip(source),after=readZip(bytes);
  expect(reopened.readCell('Input','A1')).toMatchObject({kind:'number',value:10});
  expect(reopened.readCell('Calc','A1')).toMatchObject({kind:'formula',formula:'Input!A1*2',cached:null});
  expect(reopened.readCell('Calc','B1')).toMatchObject({kind:'formula',formula:'A1+1',cached:null});
  expect(reopened.readCell('Calc','C1')).toMatchObject({kind:'formula',formula:'42',cached:42});
  expect(savedFormula(source,'xl/worksheets/sheet2.xml','A1')).toEqual({formula:'Input!A1*2',dataOnly:'2'});
  expect(savedFormula(source,'xl/worksheets/sheet2.xml','B1')).toEqual({formula:'A1+1',dataOnly:'3'});
  expect(savedFormula(bytes,'xl/worksheets/sheet2.xml','A1')).toEqual({formula:'Input!A1*2',dataOnly:null});
  expect(savedFormula(bytes,'xl/worksheets/sheet2.xml','B1')).toEqual({formula:'A1+1',dataOnly:null});
  expect(savedFormula(bytes,'xl/worksheets/sheet2.xml','C1')).toEqual({formula:'42',dataOnly:'42'});
  for(const name of before.keys())if(!allowed.has(name)&&name!==chain)expect(after.get(name)).toEqual(before.get(name));
  expect([...after.keys()].sort()).toEqual([...before.keys()].filter(name=>name!==chain).sort());
  expect(after.has(chain)).toBe(false);expect(decode.decode(after.get(rel)!)).not.toContain(relationship);
  expect(decode.decode(after.get(types)!)).not.toContain(contentType);
  expect(decode.decode(after.get('xl/workbook.xml')!)).toContain('forceFullCalc="1"');
  expect(decode.decode(after.get('xl/workbook.xml')!)).toContain('fullCalcOnLoad="1"');
  const pkg=await OpcPackage.open(bytes);for(const name of pkg.names())pkg.get(name);
  expect(pkg.relationships('xl/workbook.xml').filter(r=>r.type===relationship)).toHaveLength(0);
 }finally{await rm(root,{recursive:true,force:true});}
});

test('opt-in incomplete or ambiguous chain graphs refuse atomically',async()=>{
 const source=await ownedChainSource();
 for(const [part,find,replacement] of [
  ['xl/worksheets/sheet2.xml','Input!A1*2','SUM(UnknownName)'],
  ['xl/worksheets/sheet2.xml','<s:f>42</s:f>','<s:f>C1</s:f>'],
  ['xl/worksheets/sheet2.xml','Input!A1*2','Input!Z9*2'],
  ['xl/worksheets/sheet2.xml','<s:f>42</s:f>','<s:f>A1:C1</s:f>'],
  [rel,'Target="chains/order.xml"','TargetMode="External" Target="chains/order.xml"'],
  [types,contentType,'application/xml'],
  [chain,'i="2"','i="99"'],
  [chain,'<c r="B1" i="2"/>','<c r="B1" i="2"/><c r="B1" i="2"/>'],
  ['xl/workbook.xml','</s:workbook>','<s:definedNames><s:definedName name="Invisible">Input!A1</s:definedName></s:definedNames></s:workbook>'],
  ['xl/worksheets/sheet2.xml','<s:f>42</s:f>','<s:f t="shared">42</s:f>'],
  ['xl/worksheets/sheet2.xml','<s:f>Input!A1*2</s:f>','<s:f>B1*2</s:f>'],
  ['xl/worksheets/sheet2.xml','<s:f>42</s:f>','<s:f>C1</s:f>'],
 ] as const){
  const parts=readZip(source),xml=decode.decode(parts.get(part)!);expect(xml).toContain(find);
  parts.set(part,encode.encode(xml.replace(find,replacement)));
  const changed=writeZip(parts),wb=await Workbook.open(changed);
  expect(()=>wb.worksheet('Input').setCellValue('A1',10,{calculationPolicy:'invalidate-dependent-formula-caches'})).toThrow(expect.objectContaining({code:'xlsx-calculation-chain-unsupported'}));
  expect(wb.toBytes()).toEqual(changed);
 }
});

test('direct API refuses an unrecognised policy without touching the workbook',async()=>{
 const source=await ownedChainSource(),wb=await Workbook.open(source);
 expect(()=>wb.worksheet('Input').setCellValue('A1',10,{calculationPolicy:'invalid' as 'invalidate-dependent-formula-caches'})).toThrow(expect.objectContaining({code:'xlsx-calculation-chain-unsupported'}));
 expect(wb.toBytes()).toEqual(source);
});

test('unsupported dependency graph preserves source and existing output in workflow',async()=>{
 const parts=readZip(await ownedChainSource()),sheet='xl/worksheets/sheet2.xml';
 parts.set(sheet,encode.encode(decode.decode(parts.get(sheet)!).replace('Input!A1*2','Input!Z9*2')));
 const source=writeZip(parts),root=await mkdtemp(join(tmpdir(),'bun-chain-refusal-'));
 try{
  const input=join(root,'input.xlsx'),output=join(root,'out.xlsx'),oldOutput=encode.encode('existing output');
  await Bun.write(input,source);await Bun.write(output,'existing output');
  const result=await patchOffice({source:input,output,mode:'safe',calculationPolicy:'invalidate-dependent-formula-caches',changes:[{target:'Input!A1',value:10}]});
  expect(result.status).toBe('refused');expect(result.error?.code).toBe('xlsx-calculation-chain-unsupported');expect(result.committedChanges).toBe(0);
  expect([...await Bun.file(input).bytes()]).toEqual([...source]);
  expect([...await Bun.file(output).bytes()]).toEqual([...oldOutput]);
 }finally{await rm(root,{recursive:true,force:true});}
});
