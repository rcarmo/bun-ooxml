import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';
import {addPart,addRelationship} from '../../src/opc/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));
async function fixture(pr='',change=(s:string)=>s){
 const d=Document.create();const t=d.addTable(2,2);t.cell(0,0).text='Header';t.cell(1,1).text='Body';
 const pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),change(pkg.text(pkg.mainPart()).replace('<w:tr>','<w:tr>'+pr)));addPart(pkg,'custom/payload.bin',new Uint8Array([0,255,9]),'application/octet-stream');return Document.open(pkg.toBytes());
}

test('direct row header toggles preserve text, dimensions and other members after path save/reopen',async()=>{
 const d=await fixture(),table=d.tables[0]!,cell=table.cell(0,0),paragraph=d.paragraphs[0]!,before=d.package.parts;
 expect(table.isRowHeader(0)).toBe(false);expect(table.isRowHeader(1)).toBe(false);expect(table.setRowHeader(0,true)).toEqual({changed:1});expect(table.isRowHeader(0)).toBe(true);expect(xml(d)).toContain('w:val="on"');expect(()=>cell.text).toThrow();expect(()=>paragraph.text).toThrow();
 expect([table.rows,table.columns]).toEqual([2,2]);expect(table.cell(0,0).text).toBe('Header');expect(table.cell(1,1).text).toBe('Body');
 const root=await mkdtemp(join(tmpdir(),'row-header-'));try{const path=join(root,'table.docx');await d.save(path);const reopened=await Document.open(path);expect(reopened.tables[0]!.isRowHeader(0)).toBe(true);expect(reopened.tables[0]!.isRowHeader(1)).toBe(false);expect(reopened.tables[0]!.cell(1,1).text).toBe('Body');for(const[n,b]of before)if(n!=='word/document.xml')expect(reopened.package.get(n)).toEqual(b);}finally{await rm(root,{recursive:true,force:true});}
 expect(table.setRowHeader(0,false)).toEqual({changed:1});expect(table.isRowHeader(0)).toBe(false);expect(xml(d)).toContain('w:val="off"');expect(table.setRowHeader(0,null)).toEqual({changed:1});expect(xml(d)).not.toContain('tblHeader');
});

test('all on/off lexical values read correctly and validated no-ops retain exact XML and handles',async()=>{
 for(const[value,enabled]of [['',true],['true',true],['1',true],['on',true],['false',false],['0',false],['off',false]] as const){const d=await fixture(`<w:trPr><w:tblHeader${value?` w:val='${value}'`:''}/></w:trPr>`),t=d.tables[0]!,cell=t.cell(0,0),before=d.package.toBytes();expect(t.isRowHeader(0)).toBe(enabled);expect(t.setRowHeader(0,enabled)).toEqual({changed:0});expect(d.package.toBytes()).toEqual(before);expect(cell.text).toBe('Header');}
 const d=await fixture(),t=d.tables[0]!,before=d.package.toBytes();expect(t.setRowHeader(0,null)).toEqual({changed:0});expect(d.package.toBytes()).toEqual(before);expect(t.setRowHeader(0,false)).toEqual({changed:1});
});

test('self-closing properties and unordered trPrBase metadata retain exact sibling fragments',async()=>{
 for(const pr of ['<w:trPr/>','<w:trPr><w:jc w:val="center"/><w:trHeight w:val="400" w:hRule="atLeast"/><w:cantSplit/></w:trPr>']){const d=await fixture(pr),before=xml(d),t=d.tables[0]!;t.setRowHeader(0,true);const after=xml(d);expect(after.replace(/<w:tblHeader[^>]*\/>/,'').replace('<w:trPr></w:trPr>','<w:trPr/>')).toBe(before);t.setRowHeader(0,null);expect(xml(d).replace('<w:trPr></w:trPr>','<w:trPr/>')).toBe(before);}
 // Direct markers need not be contiguous; this API does not predict pagination.
 const d=await fixture();d.tables[0]!.setRowHeader(1,true);expect(d.tables[0]!.isRowHeader(0)).toBe(false);expect(d.tables[0]!.isRowHeader(1)).toBe(true);
});

test('invalid coordinates and non-Boolean input refuse atomically',async()=>{
 const d=await fixture(),t=d.tables[0]!,before=d.package.toBytes();for(const row of [-1,2,0.5,NaN,Infinity,'0' as unknown as number]){expect(()=>t.isRowHeader(row)).toThrow();expect(()=>t.setRowHeader(row,true)).toThrow();}
 for(const value of [undefined,1,'true',{},[],new Boolean(true)])expect(()=>t.setRowHeader(0,value as boolean)).toThrow();expect(d.package.toBytes()).toEqual(before);
});

test('duplicate, decorated, wrong-namespace, revised and lexical markers refuse even removal or no-op',async()=>{
 for(const pr of ['<w:tblHeader w:val="yes"/>','<w:tblHeader val="1"/>','<w:tblHeader xmlns:x="urn:x" x:val="1"/>','<w:tblHeader w:val="true"><w:bad/></w:tblHeader>','<w:tblHeader/><w:tblHeader/>','<w:tblHeader/><!--barrier-->','<w:ins/>','<w:trPrChange/>','<w:unknown/>','<w:tblHeader><?x y?></w:tblHeader>','<w:cantSplit/><w:cantSplit/>']){
  const d=await fixture('<w:trPr>'+pr+'</w:trPr>'),t=d.tables[0]!,before=d.package.toBytes();expect(()=>t.isRowHeader(0)).toThrow();for(const val of [true,false,null])expect(()=>t.setRowHeader(0,val)).toThrow();expect(d.package.toBytes()).toEqual(before);
 }
});

test('unsupported row/table topology, revision and merged-cell contexts refuse before mutation',async()=>{
 for(const change of [
  (x:string)=>x.replace('<w:tr>','<w:tr><!--barrier-->'),
  (x:string)=>x.replace('<w:tr>','<w:tr><w:trPr/><w:trPr/>'),
  (x:string)=>x.replace('</w:tr>','<w:trPr/></w:tr>'),
  (x:string)=>x.replace('</w:tbl>','<w:tblPr/></w:tbl>'),
  (x:string)=>x.replace(/<w:tblGrid>[^]*?<\/w:tblGrid>/,''),
  (x:string)=>x.replace('<w:tcPr>','<w:tcPr><w:vMerge/>'),
  (x:string)=>x.replace(/<w:p(?:\s[^>]*)?>/,'$&<w:ins/>'),
  (x:string)=>x.replace('<w:tblPr>','<w:tblPr><w:tblPrChange/>'),
 ]){const d=await fixture('',change),before=d.package.toBytes();expect(()=>d.tables[0]!.setRowHeader(0,true),change.toString()).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('protection and external XML changes refuse including same-value requests',async()=>{
 const d=await fixture('<w:trPr><w:tblHeader/></w:trPr>'),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(pkg.toBytes()),before=locked.package.toBytes();expect(()=>locked.tables[0]!.setRowHeader(0,true)).toThrow();expect(locked.package.toBytes()).toEqual(before);
 const t=d.tables[0]!;d.package.setPart('word/document.xml',new TextEncoder().encode(xml(d).replace('Header','Changed')));const changed=d.package.toBytes();expect(()=>t.setRowHeader(0,true)).toThrow();expect(()=>t.isRowHeader(0)).toThrow();expect(d.package.toBytes()).toEqual(changed);
 const fresh=await fixture(),stale=fresh.tables[0]!;fresh.addTable(1,1);expect(()=>stale.setRowHeader(0,true)).toThrow(expect.objectContaining({code:'docx-stale-table'}));
});

test('package write and serialization failure roll back bytes and retain paragraph/cell/table handles',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=await fixture(),t=d.tables[0]!,c=t.cell(0,0),p=d.paragraphs[0]!,before=d.package.toBytes(),pkg=(d as unknown as {opcPackage:OpcPackage}).opcPackage,original=pkg[stage].bind(pkg);
  if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialize');};
  try{expect(()=>t.setRowHeader(0,true)).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}
  expect(d.package.toBytes()).toEqual(before);expect(c.text).toBe('Header');expect(p.text).toBe('Header');expect(t.isRowHeader(0)).toBe(false);
 }
});

test('alias and conflicting lexical w bindings plus UTF-16 save retain expanded marker identity',async()=>{
 const d=await fixture('',x=>x.replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('<q:tr>','<q:tr xmlns:w="urn:foreign">')),pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()).replace('UTF-8','UTF-16');const bytes=new Uint8Array(2+text.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);
 const loaded=await Document.open(pkg.toBytes());loaded.tables[0]!.setRowHeader(0,true);const reopened=await Document.open(loaded.package.toBytes());expect(reopened.tables[0]!.isRowHeader(0)).toBe(true);expect([...reopened.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(reopened.package.get(n)).toEqual(pkg.get(n));
});

test('shared header getter executes and false readback or initial-state predicates fail assertions',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/row-header.ts');const path='workflows/docx/tables.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(3)}});
 const good=await executeAcceptance(inv(source),bindings,'row-header-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);
 for(const text of [source.replace('the first result is false and the second is true','the first result is false and the second is false'),source.replace('the first result is false and the second is true','the first result is true and the second is true')]){const bad=await executeAcceptance(inv(text),bindings,'row-header-predicate');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 const corrupt=bindings.map(b=>b.pattern.test('IsHeader is read, SetHeader true is applied and IsHeader is read again')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);const s=c.state as {document:Document;headerAfter:boolean};s.document.tables[0]!.setRowHeader(0,false);s.headerAfter=s.document.tables[0]!.isRowHeader(0);}}:b);
 const bad=await executeAcceptance(inv(source),corrupt,'row-header-mutation');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);
});

test('output bound refuses row header insertion without expiring live handles',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes()),main=pkg.mainPart(),text=pkg.text(main),padding='x'.repeat(8*1024*1024-text.length-9);pkg.set(main,text.replace('</w:body>','<!--'+padding+'--></w:body>'));
 const loaded=await Document.open(pkg.toBytes()),table=loaded.tables[0]!,cell=table.cell(0,0),before=loaded.package.toBytes();expect(()=>table.setRowHeader(0,true)).toThrow(expect.objectContaining({code:'XML_EDIT_UNSAFE'}));expect(loaded.package.toBytes()).toEqual(before);expect(cell.text).toBe('Header');expect(table.isRowHeader(0)).toBe(false);
});
