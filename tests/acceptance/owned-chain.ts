import assert from 'node:assert/strict';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {fixturePath,F} from '../../scripts/fixture-inputs.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';
import {OpcPackage} from '../../src/opc/package.ts';
import {getContentType} from '../../src/opc/content-types.ts';
import {patchOffice} from '../../src/workflow/index.ts';
import {Workbook} from '../../src/xlsx/index.ts';
import {attribute,elements,parseXml} from '../../src/xml/index.ts';

const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const chainPart='xl/chains/order.xml';
const chainRel=`${R}/calcChain`;
const chainType='application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
const relPart='xl/_rels/workbook.xml.rels';
const typesPart='[Content_Types].xml';
const mutable=new Set(['xl/workbook.xml','xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml',relPart,typesPart]);
const encoder=new TextEncoder(),decoder=new TextDecoder();
type State={source?:Uint8Array;input?:string;output?:string;root?:string;before?:Map<string,Uint8Array>;after?:Map<string,Uint8Array>;saved?:Uint8Array;reopened?:Workbook;result?:Awaited<ReturnType<typeof patchOffice>>};
const state=(c:Record<string,unknown>)=>c.state as State;
const required=<T>(value:T|undefined):T=>{assert(value!==undefined);return value;};
const document=(parts:Map<string,Uint8Array>,path:string)=>parseXml(decoder.decode(required(parts.get(path))));
function formula(parts:Map<string,Uint8Array>,ref:string){
 const cell=elements(document(parts,'xl/worksheets/sheet2.xml').root,'c',S).find(c=>attribute(c,'r')===ref);
 assert(cell,`Missing Calc!${ref}`);
 const f=cell.children.find(e=>e.localName==='f'&&e.namespaceURI===S);
 const v=cell.children.find(e=>e.localName==='v'&&e.namespaceURI===S);
 return {text:f?.text,dataOnly:v?.text||null};
}
const roots:string[]=[];
export const cleanup=async()=>{await Promise.all(roots.splice(0).map(root=>rm(root,{recursive:true,force:true})));};
export const bindings:StepBinding[]=[
 {pattern:/^a two-sheet XLSX package with Input!A1 numeric 1 and Calc!A1 formula "Input!A1\*2" cached as 2$/,run:async c=>{
  const parts=readZip(Uint8Array.from(await Bun.file(fixturePath(F.shared.crossSheetCache)).bytes()));
  assert.equal(elements(document(parts,'xl/worksheets/sheet1.xml').root,'c',S).find(e=>attribute(e,'r')==='A1')?.children.find(e=>e.localName==='v')?.text,'1');
  assert.deepEqual(formula(parts,'A1'),{text:'Input!A1*2',dataOnly:'2'});
  const xml=decoder.decode(required(parts.get('xl/worksheets/sheet2.xml')));
  assert(xml.includes('</s:row>'));
  parts.set('xl/worksheets/sheet2.xml',encoder.encode(xml.replace('</s:row>','<s:c r="B1"><s:f>A1+1</s:f><s:v>3</s:v></s:c><s:c r="C1"><s:f>42</s:f><s:v>42</s:v></s:c></s:row>')));
  state(c).before=parts;
 }},
 {pattern:/^Calc!B1 formula "A1\+1" is cached as 3 and unrelated Calc!C1 formula "42" is cached as 42$/,run:c=>{
  const parts=required(state(c).before);assert.deepEqual(formula(parts,'B1'),{text:'A1+1',dataOnly:'3'});assert.deepEqual(formula(parts,'C1'),{text:'42',dataOnly:'42'});
 }},
 {pattern:/^the workbook alone owns a calculation-chain relationship to xl\/chains\/order\.xml$/,run:async c=>{
  const parts=required(state(c).before),xml=decoder.decode(required(parts.get(relPart)));
  assert(!xml.includes(chainRel));assert(xml.includes('</rel:Relationships>'));
  parts.set(relPart,encoder.encode(xml.replace('</rel:Relationships>',`<rel:Relationship Id="rIdOwnedCalcChain" Type="${chainRel}" Target="chains/order.xml"/></rel:Relationships>`)));
  parts.set(chainPart,encoder.encode(`<calcChain xmlns="${S}"><c r="A1" i="2"/><c r="B1" i="2"/></calcChain>`));
  const links=elements(document(parts,relPart).root,'Relationship','http://schemas.openxmlformats.org/package/2006/relationships');
  assert.deepEqual(links.filter(r=>attribute(r,'Type')===chainRel).map(r=>[attribute(r,'Id'),attribute(r,'Target')]),[['rIdOwnedCalcChain','chains/order.xml']]);
  assert(!decoder.decode(required(parts.get('_rels/.rels'))).includes(chainRel));
 }},
 {pattern:/^xl\/chains\/order\.xml has the calculation-chain content type and entries for Calc!A1 and Calc!B1$/,run:async c=>{
  const parts=required(state(c).before),xml=decoder.decode(required(parts.get(typesPart)));
  assert(xml.includes('</ct:Types>'));
  parts.set(typesPart,encoder.encode(xml.replace('</ct:Types>',`<ct:Override PartName="/${chainPart}" ContentType="${chainType}"/></ct:Types>`)));
  const pkg=OpcPackage.fromParts(parts);assert.equal(getContentType(pkg,chainPart),chainType);
  const owners=pkg.names().filter(n=>n!==typesPart&&!n.endsWith('.rels')).flatMap(n=>pkg.relationships(n).filter(r=>r.type===chainRel).map(r=>({owner:n,target:r.resolved})));
  assert.deepEqual(owners,[{owner:'xl/workbook.xml',target:chainPart}]);
  const entries=elements(document(parts,chainPart).root,'c',S).map(e=>[attribute(e,'i'),attribute(e,'r')]);assert.deepEqual(entries,[['2','A1'],['2','B1']]);
 }},
 {pattern:/^all source package member payloads and bytes are recorded$/,run:async c=>{
  const s=state(c),parts=required(s.before);s.source=writeZip(parts);s.before=readZip(s.source);
  s.root=await mkdtemp(join(tmpdir(),'bun-owned-chain-acceptance-'));roots.push(s.root);
  s.input=join(s.root,'source.xlsx');s.output=join(s.root,'result.xlsx');await Bun.write(s.input,s.source);
  assert.deepEqual([...Uint8Array.from(await Bun.file(s.input).bytes())],[...s.source]);
 }},
 {pattern:/^Input!A1 is set to numeric 10 with explicit dependent-cache invalidation and the result is saved and reopened$/,run:async c=>{
  const s=state(c);s.result=await patchOffice({source:required(s.input),output:required(s.output),mode:'safe',calculationPolicy:'invalidate-dependent-formula-caches',changes:[{target:'Input!A1',value:10}]});
  assert.equal(s.result.error,undefined);assert.equal(s.result.status,'committed');assert.equal(s.result.committedChanges,1);
  s.saved=Uint8Array.from(await Bun.file(required(s.output)).bytes());s.after=readZip(s.saved);s.reopened=await Workbook.open(s.saved);
 }},
 {pattern:/^the source package bytes remain unchanged and reopened Input!A1 is numeric 10$/,run:async c=>{
  const s=state(c);assert.deepEqual([...Uint8Array.from(await Bun.file(required(s.input)).bytes())],[...required(s.source)]);
  assert.deepEqual(required(s.reopened).readCell('Input','A1'),{kind:'number',ref:'A1',styleId:undefined,value:10});
 }},
 {pattern:/^the reopened Calc!A1 and Calc!B1 formulas are unchanged with absent or empty cached values$/,run:c=>{
  const s=state(c),book=required(s.reopened);
  for(const [ref,text] of [['A1','Input!A1*2'],['B1','A1+1']] as const){const cell=book.readCell('Calc',ref);assert(cell?.kind==='formula');assert.equal(cell.formula,text);assert.equal(cell.cached,null);}
 }},
 {pattern:/^a data-only read of those cells cannot return the old cached values 2 and 3 as current$/,run:c=>{
  const parts=required(state(c).after);assert.deepEqual(formula(parts,'A1'),{text:'Input!A1*2',dataOnly:null});assert.deepEqual(formula(parts,'B1'),{text:'A1+1',dataOnly:null});
 }},
 {pattern:/^the reopened Calc!C1 formula and cached value 42 remain unchanged$/,run:c=>{
  const s=state(c),cell=required(s.reopened).readCell('Calc','C1');assert(cell?.kind==='formula');assert.equal(cell.formula,'42');assert.equal(cell.cached,42);
  assert.deepEqual(formula(required(s.after),'C1'),formula(required(s.before),'C1'));
 }},
 {pattern:/^the workbook requests full recalculation without claiming a computed result$/,run:c=>{
  const s=state(c),xml=document(required(s.after),'xl/workbook.xml');
  const calc=elements(xml.root,'calcPr',S);assert.equal(calc.length,1);assert.equal(attribute(calc[0]!,'calcMode'),'auto');assert.equal(attribute(calc[0]!,'fullCalcOnLoad'),'1');assert.equal(attribute(calc[0]!,'forceFullCalc'),'1');
  assert.equal(required(s.result).calculationState,'recalculation-required');assert.equal(formula(required(s.after),'A1').dataOnly,null);assert.equal(formula(required(s.after),'B1').dataOnly,null);
 }},
 {pattern:/^xl\/chains\/order\.xml, its workbook relationship and its content-type override are absent$/,run:async c=>{
  const s=state(c),parts=required(s.after);assert(!parts.has(chainPart));const pkg=await OpcPackage.open(required(s.saved));
  assert.equal(pkg.relationships('xl/workbook.xml').filter(r=>r.type===chainRel).length,0);
  const overrides=elements(document(parts,typesPart).root,'Override','http://schemas.openxmlformats.org/package/2006/content-types');
  assert.equal(overrides.filter(e=>attribute(e,'PartName')===`/${chainPart}`||attribute(e,'ContentType')===chainType).length,0);
 }},
 {pattern:/^every destination relationship and content-type target resolves$/,run:async c=>{
  const pkg=await OpcPackage.open(required(state(c).saved));
  for(const name of pkg.names())if(name!==typesPart){assert(getContentType(pkg,name),`Missing content type: ${name}`);if(name.endsWith('.rels'))pkg.get(name);}
  for(const owner of ['',...pkg.names().filter(n=>n!==typesPart&&!n.endsWith('.rels'))])for(const rel of pkg.relationships(owner))if(!rel.external)assert(pkg.get(required(rel.resolved)),`Missing target: ${rel.resolved}`);
 }},
 {pattern:/^every source member payload outside the workbook, two worksheets, workbook relationships and content types remains byte-identical$/,run:c=>{
  const s=state(c),before=required(s.before),after=required(s.after);
  assert.deepEqual([...after.keys()].sort(),[...before.keys()].filter(name=>name!==chainPart).sort());
  for(const [name,bytes] of before)if(!mutable.has(name)&&name!==chainPart)assert.deepEqual(after.get(name),bytes,`Changed unrelated member ${name}`);
 }},
];
