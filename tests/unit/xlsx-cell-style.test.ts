import {test,expect} from 'bun:test';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/cell-style.ts';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';

test('cell style selection executes27 saved outcomes and atomic refusals',async()=>{
 const path='workflows/xlsx/cell-style.feature',f=parseFeature('references/fixtures-ooxml/'+path,(await Bun.file(join(fixturesRoot(),path)).text()).replace(/^@planned/m,'@implemented @bun')),rows=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(rows.length),steps:count(rows.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inv,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(27);
});

import {Workbook} from '../../src/xlsx/index.ts';
import {OpcPackage,removeRelationship,removePart} from '../../src/opc/index.ts';
import {styledWorkbook,part,S,STYLE_REL} from '../acceptance/cell-style.ts';
import {parseXml,elements} from '../../src/xml/index.ts';

test('style-only edits preserve cell payloads for scalar and formula topologies',async()=>{
 for(const payload of ['<v>17</v>','<f>A2+1</f><v>18</v>','<f t="shared" si="0" ref="A1:A2">A2+1</f><v>18</v>','<f t="array" ref="A1:B2">A2+1</f><v>18</v>']){const w=await styledWorkbook(),p=w.package;p.set(part,p.text(part).replace(/<c\b[^>]*>[\s\S]*?<\/c>/,`<c r="A1">${payload}</c>`));const q=await Workbook.open(p.toBytes()),before=q.worksheet('Sheet1').getCell('A1')!;q.worksheet('Sheet1').setCellStyle('a1',1);const saved=await Workbook.open(q.toBytes());expect(saved.worksheet('Sheet1').getCell('A1')).toEqual({...before,styleId:'1'});expect(saved.package.text(part)).toContain(payload);}
});
test('null removes direct index without styles registry and explicit zero remains distinct',async()=>{
 const w=await styledWorkbook('remove'),p=w.package,main=p.mainPart(),link=p.relationships(main).find(r=>r.type===STYLE_REL)!;removeRelationship(p,main,link.id);removePart(p,'xl/styles.xml');const q=await Workbook.open(p.toBytes());expect(q.worksheet('Sheet1').setCellStyle('A1',null)).toEqual({changed:1});expect(q.worksheet('Sheet1').getCell('A1')!.styleId).toBeUndefined();const before=q.toBytes();expect(()=>q.worksheet('Sheet1').setCellStyle('A1',0)).toThrow();expect(q.toBytes()).toEqual(before);
 const r=await styledWorkbook(),sheet=r.worksheet('Sheet1');expect(sheet.setCellStyle('A1',0)).toEqual({changed:1});expect(sheet.getCell('A1')!.styleId).toBe('0');expect(sheet.setCellStyle('A1',null)).toEqual({changed:1});expect(sheet.getCell('A1')!.styleId).toBeUndefined();
});
test('same numeric style preserves lexical index spelling and exact original archive',async()=>{
 const w=await styledWorkbook('no-op');w.package.set(part,w.package.text(part).replace('s="1"',"s = '0&#49;'"));const q=await Workbook.open(w.toBytes()),before=q.toBytes();expect(q.worksheet('Sheet1').setCellStyle('A1',1)).toEqual({changed:0});expect(q.toBytes()).toEqual(before);
});
test('style attribute scanner handles quotes, prefixes and unrelated values containing s=',async()=>{
 const w=await styledWorkbook();w.package.set(part,w.package.text(part).replace('r="A1"',`r='A1' xmlns:x="urn:custom" x:note="s='99'" s = '0'`));const q=await Workbook.open(w.toBytes()),xml=q.package.text(part);q.worksheet('Sheet1').setCellStyle('A1',1);expect(q.package.text(part)).toBe(xml.replace("s = '0'",'s = "1"'));q.worksheet('Sheet1').setCellStyle('A1',null);expect(q.package.text(part)).toBe(xml.replace(" s = '0'",''));
});
test('UTF16 LE/BE worksheet declaration, BOM and target value survive style selection',async()=>{
 for(const le of [true,false]){const w=await styledWorkbook(),xml=w.package.text(part).replace('encoding="UTF-8"','encoding="UTF-16"'),bytes=new Uint8Array(2+xml.length*2),view=new DataView(bytes.buffer);view.setUint16(0,0xfeff,le);for(let i=0;i<xml.length;i++)view.setUint16(2+i*2,xml.charCodeAt(i),le);w.package.set(part,bytes);const q=await Workbook.open(w.toBytes());q.worksheet('Sheet1').setCellStyle('A1',1);const saved=await Workbook.open(q.toBytes());expect([...saved.package.get(part)!.slice(0,2)]).toEqual(le?[255,254]:[254,255]);expect(saved.package.text(part)).toContain('encoding="UTF-16"');expect(saved.worksheet('Sheet1').getCell('A1')!.value).toBe('value π');}
});
test('serialization rollback retains earlier edits, cached values and live worksheet facades',async()=>{
 const w=await styledWorkbook();w.worksheet('Sheet1').setCellValue('B2',17);const sheet=w.worksheet('Sheet1'),before=w.toBytes(),cell=sheet.getCell('A1'),serialize=OpcPackage.prototype.toBytes;OpcPackage.prototype.toBytes=function(){throw Error('injected cell-style serialization');};try{expect(()=>sheet.setCellStyle('A1',1)).toThrow('injected cell-style serialization');}finally{OpcPackage.prototype.toBytes=serialize;}
 expect(w.toBytes()).toEqual(before);expect(sheet.getCell('A1')).toEqual(cell);expect(sheet.getCell('B2')!.value).toBe(17);sheet.setCellStyle('A1',1);sheet.setCellValue('B2',18);expect(sheet.getCell('A1')!.styleId).toBe('1');expect((await Workbook.open(w.toBytes())).worksheet('Sheet1').getCell('B2')!.value).toBe(18);
});
test('external changes to another worksheet or workbook metadata refuse cached-model style updates',async()=>{
 for(const target of ['xl/worksheets/sheet2.xml','xl/workbook.xml','_rels/.rels']){const w=await styledWorkbook(),sheet=w.worksheet('Sheet1'),old=sheet.getCell('A1');w.package.set(target,w.package.text(target)+' ');const before=w.toBytes();expect(()=>sheet.setCellStyle('A1',1)).toThrow('outside');expect(w.toBytes()).toEqual(before);expect(sheet.getCell('A1')).toEqual(old);}
});
test('custom number formats and selected base dependencies are validated',async()=>{
 const w=await styledWorkbook('missing-number-format');w.package.set('xl/styles.xml',w.package.text('xl/styles.xml').replace('<fonts','<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts><fonts'));const q=await Workbook.open(w.toBytes());expect(q.worksheet('Sheet1').setCellStyle('A1',1)).toEqual({changed:1});
 for(const replacement of ['fontId="88"','fillId="88"','borderId="88"']){const r=await styledWorkbook(),xml=r.package.text('xl/styles.xml'),old='<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';expect(xml.includes(old)).toBe(true);r.package.set('xl/styles.xml',xml.replace(old,old.replace(replacement.split('=')[0]+'="0"',replacement)));const d=await Workbook.open(r.toBytes()),before=d.toBytes();expect(()=>d.worksheet('Sheet1').setCellStyle('A1',1)).toThrow();expect(d.toBytes()).toEqual(before);}
});
test('wrong-namespace style indexes and malformed direct row placement refuse',async()=>{
 for(const [from,to]of [['r="A1"','r="A1" xmlns:q="urn:wrong" q:s="1"'],['<row r="1">','<row r="2">'],['<sheetData>','<sheetData/><sheetData>']] as const){const w=await styledWorkbook(),xml=w.package.text(part);expect(xml.includes(from)).toBe(true);w.package.set(part,xml.replace(from,to));const q=await Workbook.open(w.toBytes()),before=q.toBytes();expect(()=>q.worksheet('Sheet1').setCellStyle('A1',1)).toThrow();expect(q.toBytes()).toEqual(before);}
});
test('all index argument errors refuse without changing source bytes',async()=>{
 const w=await styledWorkbook(),before=w.toBytes();for(const index of [-1,1.5,NaN,Infinity,'1',undefined,{},true,4294967296]){expect(()=>w.worksheet('Sheet1').setCellStyle('A1',index as any)).toThrow();expect(w.toBytes()).toEqual(before);}
});
test('protected no-op and removal both refuse conservatively',async()=>{
 for(const kind of ['protected-sheet','protected-workbook']){const w=await styledWorkbook(kind),before=w.toBytes();for(const value of [null,0,1]){expect(()=>w.worksheet('Sheet1').setCellStyle('A1',value)).toThrow();expect(w.toBytes()).toEqual(before);}}
});
test('disk save and reopen retain style selection, definitions and formula cache',async()=>{
 const {mkdtemp,rm}=await import('node:fs/promises'),{tmpdir}=await import('node:os'),{join}=await import('node:path'),root=await mkdtemp(join(tmpdir(),'bun-cell-style-'));try{const w=await styledWorkbook('formula'),styles=w.package.get('xl/styles.xml')!;w.worksheet('Sheet1').setCellStyle('A1',1);const file=join(root,'styled.xlsx');await w.save(file);const q=await Workbook.open(file);expect(q.worksheet('Sheet1').getCell('A1')).toEqual({kind:'formula',ref:'A1',styleId:'1',formula:'Other!B1+1',cached:8});expect(q.package.get('xl/styles.xml')).toEqual(styles);}finally{await rm(root,{recursive:true,force:true});}
});
test('worksheet aliases to the same part refuse instead of leaving a second cached model stale',async()=>{
 const w=await styledWorkbook(),rel='xl/_rels/workbook.xml.rels';w.package.set(rel,w.package.text(rel).replace('worksheets/sheet2.xml','worksheets/sheet1.xml'));const q=await Workbook.open(w.toBytes()),before=q.toBytes();expect(()=>q.worksheet('Sheet1').setCellStyle('A1',1)).toThrow();expect(q.toBytes()).toEqual(before);expect(q.worksheet('Other').getCell('A1')!.styleId).toBeUndefined();
});

test('protected cells with an existing direct style refuse same-index and removal requests',async()=>{
 for(const kind of ['protected-sheet','protected-workbook']){const w=await styledWorkbook(kind);w.package.set(part,w.package.text(part).replace('r="A1"','r="A1" s="1"'));const q=await Workbook.open(w.toBytes()),sheet=q.worksheet('Sheet1'),before=q.toBytes();expect(sheet.getCell('A1')!.styleId).toBe('1');for(const value of [1,null]){expect(()=>sheet.setCellStyle('A1',value)).toThrow();expect(q.toBytes()).toEqual(before);expect(sheet.getCell('A1')!.styleId).toBe('1');}}
});
test('same-sheet untargeted cells, row styles and column defaults remain exact',async()=>{
 const w=await styledWorkbook();w.worksheet('Sheet1').setCellValue('B2',19);w.package.set(part,w.package.text(part).replace('<sheetData>','<cols><col min="1" max="1" style="1"/></cols><sheetData>').replace('<row r="1">','<row r="1" s="1" customFormat="1">'));const q=await Workbook.open(w.toBytes()),before=q.package.text(part),target=elements(parseXml(before),'c',S).find(n=>n.attributes.r==='B2')!,raw=before.slice(target.start,target.end);q.worksheet('Sheet1').setCellStyle('A1',0);expect(q.package.text(part)).toContain(raw);expect(q.package.text(part)).toContain('<row r="1" s="1" customFormat="1">');expect(q.package.text(part)).toContain('<col min="1" max="1" style="1"/>');q.worksheet('Sheet1').setCellStyle('A1',null);expect(q.package.text(part)).toBe(before);expect(q.worksheet('Sheet1').getCell('B2')!.value).toBe(19);
});
