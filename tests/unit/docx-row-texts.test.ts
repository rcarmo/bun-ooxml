import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';
import {addPart,addRelationship} from '../../src/opc/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(change=(s:string)=>s){const d=Document.create();const t=d.addTable(2,2);for(const[r,values]of [['A1','B1'],['A2','B2']].entries())for(const[c,value]of values.entries())t.cell(r,c).text=value;const p=await OpcPackage.open(d.package.toBytes());p.set(p.mainPart(),change(p.text(p.mainPart())));return Document.open(p.toBytes());}

test('row texts read exact cells in column order and detach results without archive mutation',async()=>{
 const d=await fixture(),t=d.tables[0]!,cell=t.cell(0,0),p=d.paragraphs[0]!,before=d.package.toBytes();const first=t.rowTexts(0);expect(first).toEqual(['A1','B1']);expect(t.rowTexts(1)).toEqual(['A2','B2']);first[0]='mutated';first.push('extra');expect(t.rowTexts(0)).toEqual(['A1','B1']);expect(d.package.toBytes()).toEqual(before);expect(cell.text).toBe('A1');expect(p.text).toBe('A1');
});

test('empty, multi-paragraph, Unicode and escaped cell values read consistently after disk save',async()=>{
 const d=await fixture(),t=d.tables[0]!;t.cell(0,0).text='';t.cell(0,1).text='  雪 & < >  \nSecond';expect(t.rowTexts(0)).toEqual(['','  雪 & < >  \nSecond']);const before=d.package.parts,dir=await mkdtemp(join(tmpdir(),'row-texts-'));try{const path=join(dir,'rows.docx');await d.save(path);const after=await Document.open(path);expect(after.tables[0]!.rowTexts(0)).toEqual(['','  雪 & < >  \nSecond']);expect(after.tables[0]!.rowTexts(1)).toEqual(['A2','B2']);for(const[n,b]of before)expect(after.package.get(n)).toEqual(b);}finally{await rm(dir,{recursive:true,force:true});}
});

test('row text inspection tracks cell edits through live tables but refuses structural staleness and external changes',async()=>{
 const d=await fixture(),t=d.tables[0]!;t.cell(0,0).text='Changed';expect(t.rowTexts(0)).toEqual(['Changed','B1']);t.setRowHeader(0,true);expect(t.rowTexts(0)).toEqual(['Changed','B1']);d.insertParagraph(0,'Before');expect(()=>t.rowTexts(0)).toThrow(expect.objectContaining({code:'docx-stale-table'}));
 const live=d.tables[0]!,xml=new TextDecoder().decode(d.package.get('word/document.xml'));d.package.setPart('word/document.xml',new TextEncoder().encode(xml.replace('Changed','External')));const before=d.package.toBytes();expect(()=>live.rowTexts(0)).toThrow(expect.objectContaining({code:'docx-stale-table'}));expect(d.package.toBytes()).toEqual(before);
});

test('invalid row coordinates refuse atomically and do not change existing cell range policy',async()=>{
 const d=await fixture(),t=d.tables[0]!,before=d.package.toBytes();for(const row of [-1,2,0.5,NaN,Infinity,'0' as unknown as number])expect(()=>t.rowTexts(row)).toThrow(RangeError);expect(()=>t.cell(-1,0)).toThrow(RangeError);expect(()=>t.cell(2,0)).toThrow(RangeError);expect(d.package.toBytes()).toEqual(before);
});

test('merged, ragged, nested or field-bearing selected cells refuse rather than return partial row text',async()=>{
 for(const change of [
  (x:string)=>x.replace('<w:tcPr>','<w:tcPr><w:vMerge/>'),
  (x:string)=>x.replace('<w:tcPr>','<w:tcPr><w:gridSpan w:val="2"/>'),
  (x:string)=>x.replace('</w:tr>','<w:tc><w:p/></w:tc></w:tr>'),
  (x:string)=>x.replace('<w:t>B1</w:t>','<w:fldChar w:fldCharType="begin"/>'),
  (x:string)=>x.replace('</w:tc>','<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc>'),
 ]){const d=await fixture(change),before=d.package.toBytes();expect(()=>d.tables[0]!.rowTexts(0)).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('read-only row text remains available under protection and does not expire handles',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(pkg.toBytes()),before=locked.package.toBytes(),t=locked.tables[0]!,c=t.cell(0,0);expect(t.rowTexts(0)).toEqual(['A1','B1']);expect(locked.package.toBytes()).toEqual(before);expect(c.text).toBe('A1');
});

test('aliased UTF-16 row inspection and byte reopen preserve all package members',async()=>{
 const d=await fixture(x=>x.replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:')),pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()).replace('UTF-8','UTF-16'),bytes=new Uint8Array(2+text.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);const source=pkg.toBytes(),loaded=await Document.open(source);expect(loaded.tables[0]!.rowTexts(0)).toEqual(['A1','B1']);expect(loaded.package.toBytes()).toEqual(source);const after=await Document.open(await loaded.save());expect(after.tables[0]!.rowTexts(1)).toEqual(['A2','B2']);for(const n of pkg.names())expect(after.package.get(n)).toEqual(pkg.get(n));
});

test('canonical four-cell and first-row predicates reject wrong cells, row order and omitted row members',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/row-texts.ts');const path='workflows/docx/tables.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(4)}});
 const good=await executeAcceptance(inv(source),bindings,'row-texts-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);
 for(let index=0;index<4;index++){const values=['A1','B1','A2','B2'];values[index]='WRONG';const altered=source.replace('the four cell text getters equal A1, B1, A2 and B2 in those positions',`the four cell text getters equal ${values[0]}, ${values[1]}, ${values[2]} and ${values[3]} in those positions`),bad=await executeAcceptance(inv(altered),bindings,'row-texts-cell-control');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 const reversed=await executeAcceptance(inv(source.replace('FirstRowText returns exactly A1 and B1','FirstRowText returns exactly B1 and A1')),bindings,'row-texts-order-control');expect(reversed.counts.cases.failed).toBe(1);expect(reversed.counts.steps.failed).toBe(1);expect(reversed.counts.steps.undefined).toBe(0);
 for(const values of [['A1'],['A1','B1','extra']]){const corrupt=bindings.map(b=>b.pattern.test('its cells are set by row to A1, B1, A2 and B2')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);(c.state as {document:Document}).document.tables[0]!.rowTexts=()=>values;}}:b);const bad=await executeAcceptance(inv(source),corrupt,'row-texts-array-control');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);}
});

test('selected-row inspection does not flatten unsupported cells in other rows or reinterpret lexical text markup',async()=>{
 const d=await fixture(x=>x.replace('<w:t>B2</w:t>','<w:fldChar w:fldCharType="begin"/>').replace('<w:t>A1</w:t>','<w:t>A<!--retained-->1</w:t>')),before=d.package.toBytes();expect(d.tables[0]!.rowTexts(0)).toEqual(['A1','B1']);expect(()=>d.tables[0]!.rowTexts(1)).toThrow();expect(d.package.toBytes()).toEqual(before);
});

test('inferred grids read supported cells but an empty unresolved row is not a successful empty result',async()=>{
 const inferred=await fixture(x=>x.replace(/<w:tblGrid>[^]*?<\/w:tblGrid>/,'')),ib=inferred.package.toBytes();expect(inferred.tables[0]!.rowTexts(0)).toEqual(['A1','B1']);expect(inferred.package.toBytes()).toEqual(ib);
 const d=Document.create();d.addTable(1,1);const pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace(/<w:tblGrid>[^]*?<\/w:tblGrid>/,'').replace(/<w:tr>[^]*?<\/w:tr>/,'<w:tr/>'));const empty=await Document.open(pkg.toBytes()),before=empty.package.toBytes();expect(()=>empty.tables[0]!.rowTexts(0)).toThrow(expect.objectContaining({code:'docx-table-unsupported'}));expect(empty.package.toBytes()).toEqual(before);
});
