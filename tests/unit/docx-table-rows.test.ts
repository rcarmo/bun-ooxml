import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';import {addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));
async function fixture(change=(s:string)=>s){const d=Document.create();d.addParagraph('Before');const t=d.addTable(2,2);for(const[r,values]of [['A','B'],['C','D']].entries())for(const[c,value]of values.entries())t.cell(r,c).text=value;d.addParagraph('After');const p=await OpcPackage.open(d.package.toBytes());p.set(p.mainPart(),change(p.text(p.mainPart())));addPart(p,'custom/opaque.bin',new Uint8Array([0,255,1]),'application/octet-stream');return Document.open(p.toBytes());}

test('append insert delete return fresh tables and retain row order after real path save',async()=>{
 const d=await fixture();let t=d.tables[0]!;const held=t;t=t.appendRow();expect(t.rows).toBe(3);expect(()=>held.rows).toThrow();expect([t.cell(2,0).text,t.cell(2,1).text]).toEqual(['','']);t=t.insertRow(1);expect(t.rows).toBe(4);expect(t.cell(2,0).text).toBe('C');t=t.deleteRow(1);expect(t.rows).toBe(3);expect(d.paragraphs.map(p=>p.text)).toEqual(['Before','A','B','C','D','','','After']);
 const dir=await mkdtemp(join(tmpdir(),'table-rows-'));try{const path=join(dir,'rows.docx');await d.save(path);const after=await Document.open(path);expect(after.tables[0]!.rows).toBe(3);expect(after.tables[0]!.columns).toBe(2);expect(after.paragraphs.map(p=>p.text)).toEqual(d.paragraphs.map(p=>p.text));}finally{await rm(dir,{recursive:true,force:true});}
});

test('new rows use explicit grid widths and preserve all old row fragments and unrelated members',async()=>{
 const d=await fixture(x=>x.replace(/<w:tblGrid>[^]*?<\/w:tblGrid>/,'<w:tblGrid><w:gridCol w:w="1800"/><w:gridCol w:w="6840"/></w:tblGrid>')),before=d.package.parts,source=xml(d),doc=parseXml(source),rows=elements(doc,'tr',W).map(n=>source.slice(n.start,n.end)),grid=elements(doc,'tblGrid',W)[0]!,gridXml=source.slice(grid.start,grid.end);
 let t=d.tables[0]!.insertRow(0);expect(t.cell(0,0).directProperties().widthTwips).toBe(1800);expect(t.cell(0,1).directProperties().widthTwips).toBe(6840);expect(t.cell(1,0).text).toBe('A');for(const row of rows)expect(xml(d)).toContain(row);expect(xml(d)).toContain(gridXml);t=t.deleteRow(0);expect(xml(d)).toBe(source);const after=await Document.open(d.package.toBytes());for(const[n,b]of before)if(n!=='word/document.xml')expect(after.package.get(n)).toEqual(b);
});

test('deleting a row removes its multiple paragraphs and leaves all other bodies and tables intact',async()=>{
 const d=await fixture(x=>x.replace('</w:tc>','<w:p><w:r><w:t>Extra</w:t></w:r></w:p></w:tc>'));d.addTable(1,1);d.tables[1]!.cell(0,0).text='Other table';let t=d.tables[0]!.deleteRow(0);expect(t.rows).toBe(1);expect(t.cell(0,0).text).toBe('C');expect(d.paragraphs.map(p=>p.text)).toEqual(['Before','C','D','After','Other table']);expect(d.tables[1]!.cell(0,0).text).toBe('Other table');t=t.insertRow(1);expect(t.rows).toBe(2);expect(d.paragraphs.map(p=>p.text)).toEqual(['Before','C','D','','','After','Other table']);
});

test('row mutations expire all held paragraph span cell and table snapshots',async()=>{
 const d=await fixture(),t=d.tables[0]!,c=t.cell(0,0),p=d.paragraphs[0]!,span=p.find('Before')[0]!;const fresh=t.insertRow(0);expect(fresh).toBe(d.tables[0]!);expect(()=>t.columns).toThrow();expect(()=>c.text).toThrow();expect(()=>p.text).toThrow();await expect(span.replace('Changed')).rejects.toThrow();
});

test('invalid indexes, last-row deletion and row limits refuse without mutation',async()=>{
 const d=await fixture(),t=d.tables[0]!,before=d.package.toBytes();for(const index of [-1,3,0.5,NaN,Infinity,'1' as unknown as number])expect(()=>t.insertRow(index)).toThrow();for(const index of [-1,2,10,0.5,NaN])expect(()=>t.deleteRow(index)).toThrow();expect(d.package.toBytes()).toEqual(before);expect(t.rows).toBe(2);
 const single=Document.create();single.addTable(1,1);const sb=single.package.toBytes();expect(()=>single.tables[0]!.deleteRow(0)).toThrow();expect(single.package.toBytes()).toEqual(sb);
 const full=Document.create();full.addTable(100,1);const fb=full.package.toBytes();expect(()=>full.tables[0]!.appendRow()).toThrow();expect(full.package.toBytes()).toEqual(fb);
});

test('merged nested revised field and lexical tables refuse insertion and deletion before mutation',async()=>{
 for(const change of [
  (x:string)=>x.replace('<w:tcPr>','<w:tcPr><w:vMerge/>'),
  (x:string)=>x.replace('<w:tcPr>','<w:tcPr><w:gridSpan w:val="2"/>'),
  (x:string)=>x.replace('<w:tr>','<w:tr><w:trPr><w:ins/></w:trPr>'),
  (x:string)=>x.replace('<w:tr>','<w:tr><!--keep-->'),
  (x:string)=>x.replace('<w:tblPr>','<w:tblPr><w:tblPrChange/>'),
  (x:string)=>x.replace('<w:tblGrid>','<w:tblGrid><w:tblGridChange/>'),
  (x:string)=>x.replace('<w:t>A</w:t>','<w:fldChar w:fldCharType="begin"/>'),
  (x:string)=>x.replace('<w:t>A</w:t>','<w:t>A<!--keep--></w:t>'),
  (x:string)=>x.replace('<w:t>A</w:t>','<w:t w:unknown="x">A</w:t>'),
  (x:string)=>x.replace('</w:tc>','<w:tbl><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc>'),
 ]){const d=await fixture(change),before=d.package.toBytes();expect(()=>d.tables[0]!.insertRow(1)).toThrow();expect(()=>d.tables[0]!.deleteRow(1)).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('missing ambiguous malformed or unsupported grid widths refuse even a deletion',async()=>{
 for(const grid of ['', '<w:tblGrid/><w:tblGrid/>','<w:tblGrid><w:gridCol/><w:gridCol w:w="4320"/></w:tblGrid>','<w:tblGrid><w:gridCol w:w="-1"/><w:gridCol w:w="4320"/></w:tblGrid>','<w:tblGrid><w:gridCol w:w="31681"/><w:gridCol w:w="4320"/></w:tblGrid>','<w:tblGrid><w:gridCol w:w="1"><w:bad/></w:gridCol><w:gridCol w:w="4320"/></w:tblGrid>']){const d=await fixture(x=>x.replace(/<w:tblGrid>[^]*?<\/w:tblGrid>/,grid)),before=d.package.toBytes();expect(()=>d.tables[0]!.appendRow()).toThrow();expect(()=>d.tables[0]!.deleteRow(0)).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('protection and external XML changes refuse row mutation while preserving current package bytes',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(pkg.toBytes()),before=locked.package.toBytes();expect(()=>locked.tables[0]!.appendRow()).toThrow();expect(()=>locked.tables[0]!.deleteRow(0)).toThrow();expect(locked.package.toBytes()).toEqual(before);
 const held=d.tables[0]!;d.package.setPart('word/document.xml',new TextEncoder().encode(xml(d).replace('Before','External')));const external=d.package.toBytes();expect(()=>held.insertRow(0)).toThrow();expect(()=>held.deleteRow(0)).toThrow();expect(d.package.toBytes()).toEqual(external);
});

test('set and serialization fault injections roll back insert and delete without expiring handles',async()=>{
 for(const stage of ['set','toBytes'] as const)for(const operation of ['insert','delete']){const d=await fixture(),t=d.tables[0]!,c=t.cell(0,0),p=d.paragraphs[0]!,before=d.package.toBytes(),pkg=(d as unknown as {opcPackage:OpcPackage}).opcPackage,original=pkg[stage].bind(pkg);if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialize');};try{expect(()=>operation==='insert'?t.insertRow(1):t.deleteRow(1)).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}expect(d.package.toBytes()).toEqual(before);expect(t.rows).toBe(2);expect(c.text).toBe('A');expect(p.text).toBe('Before');}
});

test('alias and UTF-16 row edits survive byte reopen without namespace capture',async()=>{
 const d=await fixture(x=>x.replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('<q:tbl ', '<q:tbl xmlns:w="urn:foreign" ')),pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()).replace('UTF-8','UTF-16'),bytes=new Uint8Array(2+text.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);const loaded=await Document.open(pkg.toBytes());let t=loaded.tables[0]!.insertRow(1);t.cell(1,0).text='雪';t=t.deleteRow(0);const reopened=await Document.open(await loaded.save());expect(reopened.tables[0]!.rows).toBe(2);expect(reopened.tables[0]!.cell(0,0).text).toBe('雪');expect(reopened.tables[0]!.cell(1,1).text).toBe('D');expect([...reopened.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(reopened.package.get(n)).toEqual(pkg.get(n));
});

test('inserted rows do not borrow header, height, shading or direct run formatting from neighbours',async()=>{
 const d=await fixture(x=>x.replace('<w:tr>','<w:tr><w:trPr><w:tblHeader w:val="on"/><w:trHeight w:val="400"/></w:trPr>'));let t=d.tables[0]!;t.cell(0,0).setProperties({shading:'FFFF00'});d.paragraphs[1]!.setRunFormatting({bold:true});const source=xml(d),first=elements(parseXml(source),'tr',W)[0]!,raw=source.slice(first.start,first.end);t=d.tables[0]!.insertRow(1);expect(xml(d)).toContain(raw);expect(t.isRowHeader(0)).toBe(true);expect(t.isRowHeader(1)).toBe(false);expect(t.cell(1,0).directProperties().shading).toBeNull();expect(d.paragraphs[3]!.directRunFlags()).toEqual([]);const inserted=elements(parseXml(xml(d)),'tr',W)[1]!;expect(inserted.children.some(n=>n.localName==='trPr')).toBe(false);
 const reopened=await Document.open(d.package.toBytes());expect(reopened.tables[0]!.isRowHeader(0)).toBe(true);expect(reopened.tables[0]!.isRowHeader(1)).toBe(false);expect(reopened.tables[0]!.cell(0,0).directProperties().shading).toBe('FFFF00');
});

test('section-bearing paragraphs and extension attributes refuse row deletion and insertion',async()=>{
 for(const target of ['p','tr','tc','r','section']){
  const d=await fixture(x=>{const tree=parseXml(x),cell=elements(tree,'tc',W)[0]!,p=cell.children.find(n=>n.localName==='p')!,node=target==='section'||target==='p'?p:target==='r'?p.children.find(n=>n.localName==='r')!:target==='tc'?cell:cell.parent!;
   if(target==='section')return x.slice(0,p.openEnd)+'<w:pPr><w:sectPr/></w:pPr>'+x.slice(p.openEnd);
   const at=node.start+node.name.length+1;return x.slice(0,at)+' xmlns:x="urn:x" x:tag="keep"'+x.slice(at);
  }),before=d.package.toBytes();expect(xml(d)).toContain(target==='section'?'<w:pPr><w:sectPr/></w:pPr>':'x:tag="keep"');
  expect(()=>d.tables[0]!.deleteRow(0)).toThrow();expect(()=>d.tables[0]!.appendRow()).toThrow();expect(d.package.toBytes()).toEqual(before);
 }
});

test('row insertion output bound refuses without mutating package or handles',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes()),main=pkg.mainPart(),source=pkg.text(main),padding='x'.repeat(8*1024*1024-source.length-9);pkg.set(main,source.replace('</w:body>','<!--'+padding+'--></w:body>'));const loaded=await Document.open(pkg.toBytes()),held=loaded.tables[0]!,before=loaded.package.toBytes();expect(()=>held.appendRow()).toThrow(expect.objectContaining({code:'XML_EDIT_UNSAFE'}));expect(loaded.package.toBytes()).toEqual(before);expect(held.rows).toBe(2);expect(held.cell(0,0).text).toBe('A');
});

test('canonical row-count binding catches every wrong count and fabricated successful deletion',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/table-rows.ts');const path='workflows/docx/document-model.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(4)}});
 const good=await executeAcceptance(inv(source),bindings,'table-rows-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);
 for(const values of ['two, four and three','three, three and three','three, four and two']){const bad=await executeAcceptance(inv(source.replace('row counts after each step are three, four and three respectively','row counts after each step are '+values+' respectively')),bindings,'row-count-control');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 const corrupt=bindings.map(b=>b.pattern.test('row counts after each step are three, four and three respectively')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);const s=c.state as {document:Document};s.document.tables[0]!.deleteRow=()=>s.document.tables[0]!;}}:b);const bad=await executeAcceptance(inv(source),corrupt,'row-delete-control');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);
});
