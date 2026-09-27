import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';import {addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));
async function fixture(change=(s:string)=>s){const d=Document.create();d.addParagraph('Untouched');d.addTable(1,1);d.tables[0]!.cell(0,0).text='Cell';const p=await OpcPackage.open(d.package.toBytes());p.set(p.mainPart(),change(p.text(p.mainPart())));return Document.open(p.toBytes());}

test('direct title-page and background values persist through path reopen without changing content or unrelated members',async()=>{
 const d=await fixture(),before=d.package.parts,t=d.tables[0]!,cell=t.cell(0,0),held=d.paragraphs[0]!;expect(d.getDocumentProperties()).toEqual({titlePage:null,backgroundColor:null});expect(d.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'})).toEqual({changed:2});expect(d.getDocumentProperties()).toEqual({titlePage:true,backgroundColor:'EEEEEE'});expect(()=>held.text).toThrow();expect(()=>cell.text).toThrow();expect(t.rows).toBe(1);expect(t.cell(0,0).text).toBe('Cell');
 const dir=await mkdtemp(join(tmpdir(),'document-properties-'));try{const path=join(dir,'appearance.docx');await d.save(path);const after=await Document.open(path);expect(after.getDocumentProperties()).toEqual({titlePage:true,backgroundColor:'EEEEEE'});expect(after.paragraphs.map(p=>p.text)).toEqual(['Untouched','Cell']);for(const[n,b]of before)if(n!=='word/document.xml')expect(after.package.get(n)).toEqual(b);}finally{await rm(dir,{recursive:true,force:true});}
 const tree=parseXml(xml(d));expect(tree.root.children.map(n=>n.localName)).toEqual(['background','body']);const sect=elements(tree,'sectPr',W)[0]!;expect(sect.children.map(n=>n.localName)).toEqual(['pgSz','pgMar','cols','titlePg','docGrid']);
});

test('null removal, explicit false and lexical same-value no-ops keep distinct direct states',async()=>{
 const d=await fixture(x=>x.replace('<w:body>','<w:background w:color=\'aAbBcC\'/><w:body>').replace('<w:docGrid','<w:titlePg w:val=\'off\'/><w:docGrid')),before=d.package.toBytes(),held=d.paragraphs[0]!;expect(d.getDocumentProperties()).toEqual({titlePage:false,backgroundColor:'aAbBcC'});expect(d.setDocumentProperties({titlePage:false,backgroundColor:'aAbBcC'})).toEqual({changed:0});expect(d.package.toBytes()).toEqual(before);expect(held.text).toBe('Untouched');const values=d.getDocumentProperties();values.titlePage=true;expect(d.getDocumentProperties().titlePage).toBe(false);
 expect(d.setDocumentProperties({titlePage:null,backgroundColor:null})).toEqual({changed:2});expect(d.getDocumentProperties()).toEqual({titlePage:null,backgroundColor:null});expect(xml(d)).not.toContain('titlePg');expect(xml(d)).not.toContain('background');expect(d.setDocumentProperties({titlePage:false})).toEqual({changed:1});expect(d.getDocumentProperties().titlePage).toBe(false);expect(d.setDocumentProperties({backgroundColor:'auto'})).toEqual({changed:1});expect(d.getDocumentProperties().backgroundColor).toBe('auto');
});

test('all valid Boolean spellings and missing val read correctly without normalizing no-ops',async()=>{
 for(const[val,value]of [['',true],['true',true],['on',true],['1',true],['false',false],['off',false],['0',false]] as const){const d=await fixture(x=>x.replace('<w:docGrid',`<w:titlePg${val?` w:val="${val}"`:''}/><w:docGrid`)),before=d.package.toBytes();expect(d.getDocumentProperties().titlePage).toBe(value);expect(d.setDocumentProperties({titlePage:value})).toEqual({changed:0});expect(d.package.toBytes()).toEqual(before);}
});

test('self-closing sections and empty direct backgrounds allow guarded insertion without geometry synthesis',async()=>{
 const d=await fixture(x=>x.replace('<w:body>','<w:background/><w:body>').replace(/<w:sectPr>[^]*?<\/w:sectPr>/,'<w:sectPr/>'));expect(d.getDocumentProperties()).toEqual({titlePage:null,backgroundColor:null});d.setDocumentProperties({titlePage:true,backgroundColor:'000000'});expect(d.getDocumentProperties()).toEqual({titlePage:true,backgroundColor:'000000'});expect(elements(parseXml(xml(d)),'pgSz',W)).toHaveLength(0);expect((await Document.open(d.package.toBytes())).getDocumentProperties().titlePage).toBe(true);
});

test('invalid plain patches, conflicting late values and XML-unsafe colors refuse atomically',async()=>{
 const d=await fixture(),before=d.package.toBytes();let called=false;const getter=Object.defineProperty({},'titlePage',{enumerable:true,get(){called=true;return true;}});for(const patch of [null,{},[],getter,Object.create({titlePage:true}),{unknown:1},{titlePage:'true'},{titlePage:1},{backgroundColor:'#EEEEEE'},{backgroundColor:'red'},{backgroundColor:'EEEEEE\n'},{titlePage:true,backgroundColor:'bad'},{[Symbol('x')]:true}]){expect(()=>d.setDocumentProperties(patch as any)).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(called).toBe(false);
});

test('multiple sections, absent/final-order ambiguity and lexical or revised section metadata refuse unchanged',async()=>{
 for(const change of [
  (x:string)=>x.replace('<w:body>','<w:body><w:p><w:pPr><w:sectPr/></w:pPr></w:p>'),
  (x:string)=>x.replace(/<w:sectPr>[^]*?<\/w:sectPr>/,''),
  (x:string)=>x.replace('</w:body>','<w:sectPr/></w:body>'),
  (x:string)=>x.replace('</w:body>','<w:p/></w:body>'),
  (x:string)=>x.replace('<w:sectPr>','<w:sectPr><!--keep-->'),
  (x:string)=>x.replace('</w:sectPr>','<w:sectPrChange/></w:sectPr>'),
  (x:string)=>x.replace('<w:docGrid','<w:titlePg/><w:titlePg/><w:docGrid'),
  (x:string)=>x.replace('<w:pgSz','<w:titlePg/><w:pgSz'),
  (x:string)=>x.replace('<w:sectPr>','<w:sectPr><w:footnotePr><w:sectPrChange/></w:footnotePr>'),
 ]){const d=await fixture(change),before=d.package.toBytes();expect(()=>d.getDocumentProperties()).toThrow();expect(()=>d.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'})).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('themed drawing duplicate decorated or malformed background/title metadata refuse even removal',async()=>{
 for(const metadata of ['<w:background w:color="EEEEEE" w:themeColor="accent1"/>','<w:background><w:drawing/></w:background>','<w:background w:color="bad"/>','<w:background color="EEEEEE"/>','<w:background/><w:background/>','<w:background><!--keep--></w:background>']){const d=await fixture(x=>x.replace('<w:body>',metadata+'<w:body>')),before=d.package.toBytes();expect(()=>d.getDocumentProperties()).toThrow();expect(()=>d.setDocumentProperties({backgroundColor:null})).toThrow();expect(d.package.toBytes()).toEqual(before);}
 for(const marker of ['<w:titlePg w:val="yes"/>','<w:titlePg val="1"/>','<w:titlePg><w:bad/></w:titlePg>','<w:titlePg><?keep value?></w:titlePg>']){const d=await fixture(x=>x.replace('<w:docGrid',marker+'<w:docGrid')),before=d.package.toBytes();expect(()=>d.setDocumentProperties({titlePage:null})).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('root/body ambiguity and late background location refuse without destructive repair',async()=>{
 for(const change of [(x:string)=>x.replace('</w:document>','<w:background/></w:document>'),(x:string)=>x.replace('</w:document>','<w:body/></w:document>'),(x:string)=>x.replace('<w:body>','<!--keep--><w:body>'),(x:string)=>x.replace('<w:document ','<q:document xmlns:q="urn:foreign" ').replace('</w:document>','</q:document>')]){const d=await fixture(change),before=d.package.toBytes();expect(()=>d.setDocumentProperties({backgroundColor:'EEEEEE'})).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('existing geometry, header relationships, body XML and sibling metadata keep their original spelling',async()=>{
 const base=await fixture(),pkg=await OpcPackage.open(base.package.toBytes());addPart(pkg,'word/header1.xml',`<w:hdr xmlns:w="${W}"><w:p/></w:hdr>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header','header1.xml',{id:'rIdHeader'});const header='<w:headerReference xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" w:type="default" r:id="rIdHeader"/>';pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('<w:sectPr>','<w:sectPr>'+header));const d=await Document.open(pkg.toBytes()),before=d.package.parts,source=xml(d);d.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'});const next=xml(d);expect(next).toContain(header);expect(next.replace(/<w:background[^>]*\/>/,'').replace(/<w:titlePg[^>]*\/>/,'')).toBe(source);for(const[n,b]of before)if(n!=='word/document.xml')expect(d.package.get(n)).toEqual(b);
});

test('protected and externally changed document properties refuse edits including no-ops',async()=>{
 const d=await fixture(),p=await OpcPackage.open(d.package.toBytes());addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(p.toBytes()),before=locked.package.toBytes();expect(locked.getDocumentProperties().titlePage).toBeNull();expect(()=>locked.setDocumentProperties({titlePage:null})).toThrow();expect(locked.package.toBytes()).toEqual(before);
 d.package.setPart('word/document.xml',new TextEncoder().encode(xml(d).replace('Untouched','External')));const changed=d.package.toBytes();expect(()=>d.getDocumentProperties()).toThrow();expect(()=>d.setDocumentProperties({titlePage:true})).toThrow();expect(d.package.toBytes()).toEqual(changed);
});

test('write and serialization faults roll back both property edits and keep snapshot handles',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=await fixture(),held=d.paragraphs[0]!,cell=d.tables[0]!.cell(0,0),before=d.package.toBytes(),pkg=(d as any).opcPackage as OpcPackage,original=pkg[stage].bind(pkg);if(stage==='set')pkg.set=(n,v)=>{original(n as never,v as never);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialize');};try{expect(()=>d.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'})).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}expect(d.package.toBytes()).toEqual(before);expect(held.text).toBe('Untouched');expect(cell.text).toBe('Cell');expect(d.getDocumentProperties()).toEqual({titlePage:null,backgroundColor:null});}
});

test('aliased UTF16 with conflicting lexical w binding retains both properties after byte reopen',async()=>{
 const d=await fixture(x=>x.replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('<q:body>','<q:body xmlns:w="urn:foreign">')),p=await OpcPackage.open(d.package.toBytes()),source=p.text(p.mainPart()).replace('UTF-8','UTF-16'),bytes=new Uint8Array(2+source.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<source.length;i++)view.setUint16(2+i*2,source.charCodeAt(i),true);p.set(p.mainPart(),bytes);const doc=await Document.open(p.toBytes());doc.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'});const after=await Document.open(await doc.save());expect(after.getDocumentProperties()).toEqual({titlePage:true,backgroundColor:'EEEEEE'});expect([...after.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);for(const n of p.names())if(n!==p.mainPart())expect(after.package.get(n)).toEqual(p.get(n));
});

test('canonical title/background getter case detects each incorrect or corrupted direct value',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/document-properties.ts');const path='workflows/docx/document-model.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(3)}});
 const good=await executeAcceptance(inv(source),bindings,'document-properties-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);
 for(const altered of [source.replace('TitlePage getter is true','TitlePage getter is false'),source.replace('BackgroundColor getter equals EEEEEE','BackgroundColor getter equals 000000')]){const bad=await executeAcceptance(inv(altered),bindings,'document-properties-expected');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
 for(const patch of [{titlePage:false},{backgroundColor:'000000'}]){const corrupted=bindings.map(b=>b.pattern.test('TitlePage is set true on that section and BackgroundColor to EEEEEE')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);(c.state as {document:Document}).document.setDocumentProperties(patch);}}:b);const bad=await executeAcceptance(inv(source),corrupted,'document-properties-corrupt');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);}
});

test('complete XML output bounds roll back the whole two-field patch',async()=>{
 const d=await fixture(),source=xml(d),padding=' '.repeat(8*1024*1024-source.length-2);d.package.setPart('word/document.xml',new TextEncoder().encode(source.replace('</w:body>',padding+'</w:body>')));const loaded=await Document.open(d.package.toBytes()),before=loaded.package.toBytes(),held=loaded.paragraphs[0]!;expect(()=>loaded.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'})).toThrow(expect.objectContaining({code:'XML_EDIT_UNSAFE'}));expect(loaded.package.toBytes()).toEqual(before);expect(held.text).toBe('Untouched');
});

test('changed UTF8-BOM main XML retains encoding marker and untouched body fragments after reopen',async()=>{
 const d=await fixture(),p=await OpcPackage.open(d.package.toBytes()),before=p.text(p.mainPart()),raw=p.get(p.mainPart())!;p.set(p.mainPart(),new Uint8Array([239,187,191,...raw]));const loaded=await Document.open(p.toBytes());loaded.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'});const after=await Document.open(loaded.package.toBytes()),bytes=after.package.get('word/document.xml')!;expect([...bytes.slice(0,3)]).toEqual([239,187,191]);expect(after.getDocumentProperties()).toEqual({titlePage:true,backgroundColor:'EEEEEE'});expect(new TextDecoder().decode(bytes).replace(/<w:background[^>]*\/>/,'').replace(/<w:titlePg[^>]*\/>/,'')).toBe(before);
 // The shared property commit must retain that marker on a subsequent direct edit too.
 after.paragraphs[0]!.setProperties({alignment:'center'});expect([...after.package.get('word/document.xml')!.slice(0,3)]).toEqual([239,187,191]);
});
