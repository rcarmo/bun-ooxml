import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';import {addPart,addRelationship} from '../../src/opc/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));
function make(){const d=Document.create();d.addParagraph('Introduction').setProperties({outlineLevel:0});d.addParagraph('<Customer Name>');const t=d.addTable(2,2);t.cell(0,0).text='Role';t.cell(0,1).text='Count';t.cell(1,0).text='Architect';t.cell(1,1).text='1';return d;}

test('body map reports exact saved heading table placeholder and anchor counts without changing file or package bytes',async()=>{
 const d=make(),root=await mkdtemp(join(tmpdir(),'body-map-'));try{const path=join(root,'map.docx');await d.save(path);const original=await Bun.file(path).bytes(),loaded=await Document.open(path),m=loaded.inspectBodyMap();expect(m.counts).toEqual({sections:1,tables:1,placeholders:1,anchors:2});expect(m.anchors.map(a=>a.text)).toEqual(['Introduction','<Customer Name>']);expect(m.placeholders).toEqual([{bodyIndex:1,start:0,end:15,text:'<Customer Name>',name:'Customer Name'}]);expect(loaded.tables[0]!.rowTexts(0)).toEqual(['Role','Count']);expect(loaded.tables[0]!.rowTexts(1)).toEqual(['Architect','1']);expect(loaded.package.toBytes()).toEqual(original);expect(await Bun.file(path).bytes()).toEqual(original);}finally{await rm(root,{recursive:true,force:true});}
});

test('empty documents and nonheading body paragraphs never acquire synthetic section counts',()=>{
 const d=Document.create();expect(d.inspectBodyMap().counts).toEqual({sections:0,tables:0,placeholders:0,anchors:0});d.addParagraph('Ordinary');d.addParagraph('Explicit body').setProperties({outlineLevel:9});d.addParagraphStyle('Heading1',{name:'Heading1'});d.addParagraph('Style only',{style:'Heading1'});const m=d.inspectBodyMap();expect(m.counts).toEqual({sections:0,tables:0,placeholders:0,anchors:3});expect(m.anchors.every(a=>a.type==='paragraph')).toBe(true);
});

test('all direct heading levels count markers including repeated titles rather than Word section breaks',()=>{
 const d=Document.create();for(let i=0;i<9;i++)d.addParagraph('Repeated').setProperties({outlineLevel:i});d.addParagraph('Nonheading').setProperties({outlineLevel:9});const m=d.inspectBodyMap();expect(m.counts.sections).toBe(9);expect(m.counts.anchors).toBe(10);expect(m.anchors.map(a=>a.bodyIndex)).toEqual([0,1,2,3,4,5,6,7,8,9]);expect(xml(d).split('<w:sectPr>').length-1).toBe(1);
});

test('placeholder scanning preserves repeated text whitespace and UTF16 offsets across direct runs',()=>{
 const d=Document.create();d.addParagraph('😀 <Name> <Name> < spaced name >').appendRun(' <Across').appendRun(' runs>');const m=d.inspectBodyMap();expect(m.counts.placeholders).toBe(4);expect(m.placeholders.map(p=>[p.start,p.end,p.text,p.name])).toEqual([[3,9,'<Name>','Name'],[10,16,'<Name>','Name'],[17,32,'< spaced name >',' spaced name '],[33,46,'<Across runs>','Across runs']]);const text=m.anchors[0]!.text;for(const p of m.placeholders)expect(text.slice(p.start,p.end)).toBe(p.text);
});

test('placeholder grammar ignores empty whitespace-only nested oversized and unterminated spans',()=>{
 const d=Document.create();d.addParagraph('<> <   > <outer<inner>> <'+ 'x'.repeat(257)+'> <open');expect(d.inspectBodyMap().placeholders).toEqual([]);d.addParagraph('<'+ 'x'.repeat(256)+'>');expect(d.inspectBodyMap().placeholders).toHaveLength(1);expect(d.inspectBodyMap().placeholders[0]!.name).toHaveLength(256);
});

test('map table counts are top-level only and exclude table-cell text and nested tables from discovery',async()=>{
 const d=make();d.tables[0]!.cell(0,0).text='<Cell token>';d.addParagraph('After <Tail>');const m=d.inspectBodyMap();expect(m.counts).toEqual({sections:1,tables:1,placeholders:2,anchors:3});expect(m.placeholders.map(p=>[p.bodyIndex,p.name])).toEqual([[1,'Customer Name'],[3,'Tail']]);
 const pkg=await OpcPackage.open(d.package.toBytes()),source=pkg.text(pkg.mainPart());pkg.set(pkg.mainPart(),source.replace('</w:tc>','<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:tc>'));const loaded=await Document.open(pkg.toBytes()),before=loaded.package.toBytes();expect(loaded.inspectBodyMap().counts.tables).toBe(1);expect(loaded.inspectBodyMap().counts.anchors).toBe(3);expect(loaded.package.toBytes()).toEqual(before);
});

test('map records are frozen snapshots and their issued anchors retain normal ownership and staleness',()=>{
 const d=make(),m=d.inspectBodyMap();expect(Object.isFrozen(m)).toBe(true);expect(Object.isFrozen(m.counts)).toBe(true);expect(Object.isFrozen(m.placeholders)).toBe(true);expect(m.placeholders.every(Object.isFrozen)).toBe(true);expect(()=>{(m.counts as any).sections=999;}).toThrow();expect(()=>{(m.placeholders[0] as any).name='wrong';}).toThrow();expect(d.inspectBodyMap().counts.sections).toBe(1);d.insertParagraphAfter(m.anchors[0]!,'Inserted');expect(m.counts.anchors).toBe(2);expect(d.inspectBodyMap().counts.anchors).toBe(3);const before=d.package.toBytes();expect(()=>d.insertParagraphAfter(m.anchors[0]!,'stale')).toThrow(expect.objectContaining({code:'docx-stale-anchor'}));expect(d.package.toBytes()).toEqual(before);
});

test('body map refuses unsupported body paragraph topology and stale main XML without partial counts',async()=>{
 for(const change of [(s:string)=>s.replace('<w:t>Introduction</w:t>','<w:fldChar w:fldCharType="begin"/>'),(s:string)=>s.replace('w:val="0"','w:val="10"'),(s:string)=>s.replace('</w:body>','<w:sdt/></w:body>')]){const d=make(),pkg=await OpcPackage.open(d.package.toBytes()),old=pkg.text(pkg.mainPart()),next=change(old);expect(next).not.toBe(old);pkg.set(pkg.mainPart(),next);const loaded=await Document.open(pkg.toBytes()),before=loaded.package.toBytes();expect(()=>loaded.inspectBodyMap()).toThrow();expect(loaded.package.toBytes()).toEqual(before);}
 const d=make();d.package.setPart('word/document.xml',new TextEncoder().encode(xml(d)+' '));const before=d.package.toBytes();expect(()=>d.inspectBodyMap()).toThrow(expect.objectContaining({code:'docx-stale-document'}));expect(d.package.toBytes()).toEqual(before);
});

test('protection permits body-map inspection without modifying any member or document text',async()=>{
 const d=make(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const loaded=await Document.open(pkg.toBytes()),before=loaded.package.toBytes();expect(loaded.inspectBodyMap().counts).toEqual({sections:1,tables:1,placeholders:1,anchors:2});expect(loaded.paragraphs[1]!.text).toBe('<Customer Name>');expect(loaded.package.toBytes()).toEqual(before);
});

test('aliased UTF16 body maps preserve decoded placeholders and the full original archive',async()=>{
 const d=make(),pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()).replaceAll('w:','q:').replaceAll('xmlns:w=','xmlns:q=').replace('UTF-8','UTF-16'),{utf16}=await import('../fixtures/admission.ts');pkg.set(pkg.mainPart(),utf16(text));const bytes=pkg.toBytes(),loaded=await Document.open(bytes);expect(loaded.inspectBodyMap().counts).toEqual({sections:1,tables:1,placeholders:1,anchors:2});expect(loaded.inspectBodyMap().placeholders[0]!.name).toBe('Customer Name');expect(loaded.package.toBytes()).toEqual(bytes);
});

test('placeholder occurrence limit rejects excess output without invalidating existing paragraph handles',()=>{
 const d=Document.create();d.addParagraph('<x>'.repeat(10000));expect(d.inspectBodyMap().counts.placeholders).toBe(10000);d.addParagraph('<y>');const p=d.paragraphs[0]!,before=d.package.toBytes();expect(()=>d.inspectBodyMap()).toThrow(expect.objectContaining({code:'docx-body-map-limit'}));expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('<x>'.repeat(10000));
});

test('canonical saved map case checks all counts and a nonempty anchor list',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings,cleanup}=await import('../acceptance/steps.ts'),path='workflows/docx/anchor-discovery.feature',text=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const selected=selectSharedScenarios(path,text,['@id-python-word-anchor-document-map']);try{const result=await executeAcceptance({root:'.',features:[selected],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(5)}},bindings,'body-map-canonical');expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(1);expect(result.counts.cases.planned).toBe(4);expect(result.counts.steps.passed).toBe(5);}finally{await cleanup();}
});

test('canonical map predicates independently reject false counts empty anchors and changed file custody',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings,cleanup}=await import('../acceptance/steps.ts'),path='workflows/docx/anchor-discovery.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const run=async(text=source,active=bindings)=>{try{return await executeAcceptance({root:'.',features:[selectSharedScenarios(path,text,['@id-python-word-anchor-document-map'])],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(5)}},active,'map-negative');}finally{await cleanup();}};
 for(const field of ['sections','tables','placeholders','anchors','list']){const bad=bindings.map(b=>b.pattern.test('its Word document map is requested')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);const s=c.state as any;s.map=field==='list'?{...s.map,anchors:[]}:{...s.map,counts:{...s.map.counts,[field]:0}};}}:b);const result=await run(source,bad);expect(result.counts.cases.failed).toBe(1);expect(result.counts.steps.failed).toBe(1);expect(result.counts.steps.undefined).toBe(0);expect(result.counts.steps.ambiguous).toBe(0);}
 for(const [from,to]of [['counts 1 section and 1 table','counts 2 section and 1 table'],['counts 1 section and 1 table','counts 1 section and 2 table'],['at least 1 placeholder','at least 2 placeholder'],['at least 2 anchors','at least 3 anchors']]){const result=await run(source.replace(from!,to!));expect(result.counts.cases.failed).toBe(1);expect(result.counts.steps.undefined).toBe(0);}
 const given='a saved Word document has heading "Introduction", paragraph "<Customer Name>", and a 2x2 Role/Count table with Architect and 1';for(const target of ['file','package']){const bad=bindings.map(b=>b.pattern.test(given)?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);const s=c.state as {path:string;document:Document};if(target==='file')await Bun.write(s.path,'corrupt');else s.document.package.setPart('[Content_Types].xml',new TextEncoder().encode(new TextDecoder().decode(s.document.package.get('[Content_Types].xml'))+' '));expect(()=>s.document.package.toBytes()).not.toThrow();}}:b);const result=await run(source,bad);expect(result.counts.cases.failed).toBe(1);expect(result.counts.steps.undefined).toBe(0);const failed=result.features.flatMap(f=>f.scenarios.flatMap(s=>s.cases.flatMap(c=>c.steps))).filter(s=>s.status==='failed');expect(failed).toHaveLength(1);expect(failed[0]!.text).toBe('its Word document map is requested');expect(failed[0]!.error).toContain('AssertionError');}
});

test('canonical map temporary source directories are removed after positive and negative count assertions',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings,cleanup}=await import('../acceptance/steps.ts'),{lstat}=await import('node:fs/promises'),path='workflows/docx/anchor-discovery.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n}),roots:string[]=[];
 const capturing=bindings.map(b=>b.pattern.test('its Word document map is requested')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{roots.push((c.state as {root:string}).root);await b.run(c,...args);}}:b);
 for(const text of [source,source.replace('counts 1 section and 1 table','counts 2 section and 1 table')]){try{const result=await executeAcceptance({root:'.',features:[selectSharedScenarios(path,text,['@id-python-word-anchor-document-map'])],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(5)}},capturing,'map-cleanup');expect(result.counts.cases.failed).toBe(text===source?0:1);}finally{await cleanup();}}
 expect(roots).toHaveLength(2);for(const root of roots)await expect(lstat(root)).rejects.toMatchObject({code:'ENOENT'});
});
