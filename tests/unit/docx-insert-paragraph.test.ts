import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage,type AddParagraphOptions} from '../../src/index.ts';
import {addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));
const body=(text:string)=>elements(parseXml(text),'body',W)[0]!;
async function changed(change:(s:string)=>string){const d=Document.create();d.addParagraph('First');d.addParagraph('Third');const p=await OpcPackage.open(d.package.toBytes());p.set(p.mainPart(),change(p.text(p.mainPart())));return Document.open(p.toBytes());}

test('insert at start, middle and end returns correct flattened paragraph and survives path reopen',async()=>{
 const d=Document.create();expect(d.insertParagraph(0,'First').text).toBe('First');expect(d.insertParagraph(1,'Third').text).toBe('Third');const p=d.insertParagraph(1,'Second',{bold:true});expect(p.index).toBe(1);expect(p.text).toBe('Second');expect(p.directRunFlags()[0]!.bold).toBe(true);expect(d.paragraphs.map(p=>p.text)).toEqual(['First','Second','Third']);d.insertParagraph(0,'Start');d.insertParagraph(4,'End');expect(d.paragraphs.map(p=>p.text)).toEqual(['Start','First','Second','Third','End']);
 const dir=await mkdtemp(join(tmpdir(),'insert-paragraph-'));try{const path=join(dir,'body.docx');await d.save(path);expect((await Document.open(path)).paragraphs.map(p=>p.text)).toEqual(['Start','First','Second','Third','End']);}finally{await rm(dir,{recursive:true,force:true});}
});

test('body index counts tables not table-cell paragraphs; insertion preserves old fragments and unrelated members',async()=>{
 const d=Document.create();d.addParagraph('Before');const t=d.addTable(1,2);t.cell(0,0).text='A';t.cell(0,1).text='B';d.addParagraph('After');const before=d.package.parts,source=xml(d),children=body(source).children.map(n=>source.slice(n.start,n.end));
 const p=d.insertParagraph(2,'Between',{italic:true});expect(p.index).toBe(3);expect(p.text).toBe('Between');expect(d.paragraphs.map(p=>p.text)).toEqual(['Before','A','B','Between','After']);expect(d.tables[0]!.cell(0,1).text).toBe('B');expect(body(xml(d)).children.map(n=>n.localName)).toEqual(['p','tbl','p','p','sectPr']);for(const raw of children)expect(xml(d)).toContain(raw);
 const reopened=await Document.open(d.package.toBytes());expect(reopened.paragraphs[3]!.text).toBe('Between');expect([reopened.tables[0]!.rows,reopened.tables[0]!.columns]).toEqual([1,2]);for(const[n,b]of before)if(n!=='word/document.xml')expect(reopened.package.get(n)).toEqual(b);
});

test('empty insertion is structural and expires paragraphs, spans, table and cell handles',async()=>{
 const d=Document.create();d.addParagraph('Old');d.addTable(1,1);const p=d.paragraphs[0]!,span=p.find('Old')[0]!,table=d.tables[0]!,cell=table.cell(0,0);const inserted=d.insertParagraph(1,'');expect(inserted.text).toBe('');expect(d.paragraphs).toHaveLength(3);expect(()=>p.text).toThrow();await expect(span.replace('x')).rejects.toThrow();expect(()=>table.rows).toThrow();expect(()=>cell.text).toThrow();
});

test('invalid indexes, text and executable options refuse atomically without accessors',()=>{
 const d=Document.create();d.addParagraph('Old');const before=d.package.toBytes(),p=d.paragraphs[0]!;
 for(const index of [-1,2,0.5,NaN,Infinity,'0' as unknown as number])expect(()=>d.insertParagraph(index,'x')).toThrow();
 for(const text of [null,undefined,1,'\u0000','\ud800','a\nb','a\rb','a\tb','x'.repeat(1024*1024+1)])expect(()=>d.insertParagraph(0,text as string)).toThrow();
 let called=false;const getter=Object.defineProperty({},'bold',{enumerable:true,get(){called=true;return true;}});
 for(const options of [null,[],Object.create({bold:true}),getter,{bold:'yes'},{style:''},{unknown:true},{[Symbol('x')]:true}])expect(()=>d.insertParagraph(0,'x',options as AddParagraphOptions)).toThrow();expect(called).toBe(false);expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Old');
});

test('style references require existing paragraph definitions and retain styles bytes after insertion',async()=>{
 const d=Document.create();d.addParagraphStyle('Heading1',{name:'Heading 1',bold:true});const styles=d.package.get('word/styles.xml'),before=d.package.toBytes();expect(()=>d.insertParagraph(0,'bad',{style:'Missing'})).toThrow();expect(d.package.toBytes()).toEqual(before);const p=d.insertParagraph(0,'Styled',{style:'Heading1',italic:true});expect(p.styleId).toBe('Heading1');expect(p.directRunFlags()[0]!.italic).toBe(true);expect(d.package.get('word/styles.xml')).toEqual(styles);expect((await Document.open(d.package.toBytes())).paragraphs[0]!.styleId).toBe('Heading1');
});

test('unknown body children, duplicate or misplaced sections, lexical barriers and foreign roots refuse',async()=>{
 for(const change of [
  (x:string)=>x.replace('<w:body>','<w:body><!--keep-->'),
  (x:string)=>x.replace('<w:body>','<w:body><?keep value?>'),
  (x:string)=>x.replace('<w:body>','<w:body>unexpected'),
  (x:string)=>x.replace('<w:body>','<w:body><w:sdt/>'),
  (x:string)=>x.replace('<w:body>','<w:body><w:sectPr/>'),
  (x:string)=>x.replace('</w:body>','<w:p/></w:body>'),
  (x:string)=>x.replace('</w:document>','<w:body/></w:document>'),
  (x:string)=>x.replace('<w:document ', '<q:document xmlns:q="urn:foreign" ').replace('</w:document>','</q:document>'),
 ]){const d=await changed(change),before=d.package.toBytes();expect(()=>d.insertParagraph(1,'No')).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('protection and external source changes refuse insertion without clobbering package edits',async()=>{
 const d=Document.create();d.addParagraph('Old');const pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(pkg.toBytes()),before=locked.package.toBytes();expect(()=>locked.insertParagraph(0,'New')).toThrow();expect(locked.package.toBytes()).toEqual(before);
 d.package.setPart('word/document.xml',new TextEncoder().encode(xml(d).replace('Old','External')));const external=d.package.toBytes();expect(()=>d.insertParagraph(0,'New')).toThrow(expect.objectContaining({code:'docx-stale-document'}));expect(d.package.toBytes()).toEqual(external);
});

test('staged write and serialization faults preserve archive and all held handles',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=Document.create();d.addParagraph('Old');d.addTable(1,1);const p=d.paragraphs[0]!,span=p.find('Old')[0]!,t=d.tables[0]!,c=t.cell(0,0),before=d.package.toBytes(),pkg=(d as unknown as {opcPackage:OpcPackage}).opcPackage,original=pkg[stage].bind(pkg);
 if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialize');};try{expect(()=>d.insertParagraph(1,'New')).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Old');expect(t.rows).toBe(1);expect(c.text).toBe('');await span.replace('Works');expect(d.paragraphs[0]!.text).toBe('Works');}
});

test('aliased UTF-16 bodies under foreign w bindings retain insertion and original children after save',async()=>{
 const d=await changed(x=>x.replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('<q:body>','<q:body xmlns:w="urn:foreign">')),pkg=await OpcPackage.open(d.package.toBytes()),source=pkg.text(pkg.mainPart()).replace('UTF-8','UTF-16'),bytes=new Uint8Array(2+source.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<source.length;i++)view.setUint16(2+2*i,source.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);
 const loaded=await Document.open(pkg.toBytes());loaded.insertParagraph(1,'  雪 & < >  ',{bold:true});const reopened=await Document.open(await loaded.save());expect(reopened.paragraphs.map(p=>p.text)).toEqual(['First','  雪 & < >  ','Third']);expect(reopened.paragraphs[1]!.directRunFlags()[0]!.bold).toBe(true);expect([...reopened.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(reopened.package.get(n)).toEqual(pkg.get(n));
});

test('self-closing and sectionless bodies support bounded insertion without inventing sections',async()=>{
 for(const content of ['<w:body/>','<w:body></w:body>','<w:body><w:p/></w:body>']){const d=Document.create(),pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace(/<w:body>[^]*?<\/w:body>/,content));const opened=await Document.open(pkg.toBytes());expect(opened.insertParagraph(0,'New').text).toBe('New');expect(elements(parseXml(xml(opened)),'sectPr',W)).toHaveLength(0);expect((await Document.open(opened.package.toBytes())).paragraphs[0]!.text).toBe('New');}
});

test('body-order canonical case rejects incorrect counts and every reordered paragraph expectation',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/body-insertion.ts');const path='workflows/docx/paragraphs.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(4)}});
 const good=await executeAcceptance(inv(source),bindings,'body-insertion-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);
 for(const replacement of ['element counts after each operation are zero, two and three','element counts after each operation are one, one and three','element counts after each operation are one, two and two','paragraph texts in order equal Wrong, Second and Third','paragraph texts in order equal First, Wrong and Third','paragraph texts in order equal First, Second and Wrong']){const altered=source.replace(replacement.startsWith('element')?'element counts after each operation are one, two and three':'paragraph texts in order equal First, Second and Third',replacement),bad=await executeAcceptance(inv(altered),bindings,'body-insertion-predicate');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 const corrupt=bindings.map(b=>b.pattern.test('First and Third paragraphs are appended, then Second is inserted at index one')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);(c.state as {document:Document}).document.paragraphs[1]!.setText('CORRUPT');}}:b);const bad=await executeAcceptance(inv(source),corrupt,'body-insertion-corrupt');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);
});

test('complete XML output bounds refuse insertion and preserve existing handles',async()=>{
 const d=Document.create();d.addParagraph('Old');const pkg=await OpcPackage.open(d.package.toBytes()),main=pkg.mainPart(),source=pkg.text(main),padding=' '.repeat(8*1024*1024-source.length-2);pkg.set(main,source.replace('</w:body>',padding+'</w:body>'));const loaded=await Document.open(pkg.toBytes()),held=loaded.paragraphs[0]!,before=loaded.package.toBytes();expect(()=>loaded.insertParagraph(1,'x')).toThrow(expect.objectContaining({code:'XML_EDIT_UNSAFE'}));expect(loaded.package.toBytes()).toEqual(before);expect(held.text).toBe('Old');
});

test('untouched fields and earlier section properties remain exact when inserting a separate body block',async()=>{
 const field='<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>';
 const d=await changed(x=>x.replace('<w:body>','<w:body>'+field)),before=xml(d);expect(d.insertParagraph(1,'New').text).toBe('New');expect(xml(d)).toContain(field);expect(d.paragraphs[1]!.text).toBe('New');expect(xml(d).replace(/<w:p xmlns:w="[^"]+"><w:r><w:t>New<\/w:t><\/w:r><\/w:p>/,'')).toBe(before);
});
