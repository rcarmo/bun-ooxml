import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';import {addPart,addRelationship} from '../../src/opc/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));
async function fixture(change=(s:string)=>s){const d=Document.create();d.addParagraph('Introduction').setProperties({outlineLevel:0});d.addParagraph('Customer context');d.addParagraph('Delivery approach').setProperties({outlineLevel:1});d.addParagraph('Use iterative delivery');const p=await OpcPackage.open(d.package.toBytes());p.set(p.mainPart(),change(p.text(p.mainPart())));return Document.open(p.toBytes());}

test('body anchors read explicit outlines and paragraphs in saved order and filter case-insensitively without mutation',async()=>{
 const d=await fixture(),before=d.package.toBytes(),anchors=d.inspectBodyAnchors();expect(anchors.map(a=>[a.type,a.bodyIndex,a.text,a.outlineLevel])).toEqual([['section_heading',0,'Introduction',0],['paragraph',1,'Customer context',null],['section_heading',2,'Delivery approach',1],['paragraph',3,'Use iterative delivery',null]]);expect(d.inspectBodyAnchors('DeLiVeRy').map(a=>a.text)).toEqual(['Delivery approach','Use iterative delivery']);expect(d.inspectBodyAnchors('absent')).toEqual([]);expect(d.inspectBodyAnchors('')).toEqual(anchors);expect(d.package.toBytes()).toEqual(before);expect(Object.isFrozen(anchors)).toBe(true);expect(anchors.every(Object.isFrozen)).toBe(true);
});

test('insert after captured heading saves immediately adjacent text and preserves every unrelated member',async()=>{
 const d=await fixture(),anchor=d.inspectBodyAnchors('delivery').find(a=>a.text==='Delivery approach')!,before=d.package.parts,held=d.paragraphs[0]!,span=held.find('Introduction')[0]!;const inserted=d.insertParagraphAfter(anchor,'Inserted after discovered anchor',{bold:true});expect(inserted.text).toBe('Inserted after discovered anchor');expect(inserted.directRunFlags()[0]!.bold).toBe(true);expect(()=>held.text).toThrow();expect(()=>span.replace('stale')).toThrow();expect(()=>d.insertParagraphAfter(anchor,'stale')).toThrow(expect.objectContaining({code:'docx-stale-anchor'}));
 const root=await mkdtemp(join(tmpdir(),'body-anchor-'));try{const path=join(root,'anchors.docx');await d.save(path);const after=await Document.open(path);expect(after.paragraphs.map(p=>p.text)).toEqual(['Introduction','Customer context','Delivery approach','Inserted after discovered anchor','Use iterative delivery']);for(const[n,b]of before)if(n!=='word/document.xml')expect(after.package.get(n)).toEqual(b);}finally{await rm(root,{recursive:true,force:true});}
});

test('body index counts tables while excluding cell paragraphs and final section properties',async()=>{
 const d=Document.create();d.addParagraph('Before');d.addTable(1,1);d.tables[0]!.cell(0,0).text='Not a body anchor';d.addParagraph('After').setProperties({outlineLevel:8});const anchors=d.inspectBodyAnchors();expect(anchors.map(a=>[a.bodyIndex,a.text])).toEqual([[0,'Before'],[2,'After']]);const table=d.tables[0]!;d.insertParagraphAfter(anchors[0]!,'Between');expect(d.inspectBodyAnchors().map(a=>[a.bodyIndex,a.text])).toEqual([[0,'Before'],[1,'Between'],[3,'After']]);expect(d.tables[0]!.cell(0,0).text).toBe('Not a body anchor');expect(()=>table.rows).toThrow();
});

test('style names never manufacture headings and explicit nine differs from absence',async()=>{
 const d=Document.create();d.addParagraphStyle('Heading1',{name:'Heading One'});d.addParagraph('Style only',{style:'Heading1'});d.addParagraph('Explicit body').setProperties({outlineLevel:9});d.addParagraph('Direct heading').setProperties({outlineLevel:7});const rows=d.inspectBodyAnchors();expect(rows.map(a=>[a.type,a.outlineLevel])).toEqual([['paragraph',null],['paragraph',9],['section_heading',7]]);expect(d.paragraphs[0]!.styleId).toBe('Heading1');
});

test('foreign cloned forged and raw-byte-stale anchors refuse without changing either document',async()=>{
 const d=await fixture(),other=await fixture(),a=d.inspectBodyAnchors()[0]!,before=d.package.toBytes(),ob=other.package.toBytes();for(const target of [other.inspectBodyAnchors()[0]!,{...a},null]){expect(()=>d.insertParagraphAfter(target as typeof a,'bad')).toThrow(expect.objectContaining({code:'docx-stale-anchor'}));expect(d.package.toBytes()).toEqual(before);}expect(other.package.toBytes()).toEqual(ob);
 // A BOM-only external edit has identical decoded XML but different exact bytes.
 const raw=d.package.get('word/document.xml')!;d.package.setPart('word/document.xml',new Uint8Array([239,187,191,...raw]));const changed=d.package.toBytes();expect(()=>d.insertParagraphAfter(a,'bad')).toThrow(expect.objectContaining({code:'docx-stale-anchor'}));expect(d.package.toBytes()).toEqual(changed);
});

test('successful edits invalidate anchors but validated no-ops and failed inserts retain them',async()=>{
 const d=await fixture(),a=d.inspectBodyAnchors()[0]!;d.paragraphs[0]!.setProperties({outlineLevel:0});expect(()=>d.insertParagraphAfter(a,'bad\nline')).toThrow();d.insertParagraphAfter(a,'allowed');const fresh=d.inspectBodyAnchors()[0]!;d.paragraphs[1]!.setText('Changed');const before=d.package.toBytes();expect(()=>d.insertParagraphAfter(fresh,'stale')).toThrow(expect.objectContaining({code:'docx-stale-anchor'}));expect(d.package.toBytes()).toEqual(before);
});

test('protection permits discovery but refuses insertion including empty inserted text',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(pkg.toBytes()),before=locked.package.toBytes(),anchors=locked.inspectBodyAnchors();expect(anchors).toHaveLength(4);expect(()=>locked.insertParagraphAfter(anchors[0]!,'')).toThrow(expect.objectContaining({code:'docx-format-protected'}));expect(locked.package.toBytes()).toEqual(before);
});

test('unsupported or malformed body paragraphs refuse the entire discovery even when query would exclude them',async()=>{
 for(const change of [
  (s:string)=>s.replace('<w:t>Use iterative delivery</w:t>','<w:fldChar w:fldCharType="begin"/>'),
  (s:string)=>s.replace('w:val="1"','w:val="10"'),
  (s:string)=>s.replace('</w:body>','<w:sdt/></w:body>'),
  (s:string)=>s.replace('<w:body>','<w:body><!--lexical-->'),
 ]){const d=await fixture(change),before=d.package.toBytes();expect(()=>d.inspectBodyAnchors('Introduction')).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('empty documents produce no anchors and invalid queries refuse without coercion',()=>{
 const d=Document.create(),before=d.package.toBytes();expect(d.inspectBodyAnchors()).toEqual([]);for(const q of [null,1,{},'x'.repeat(4097)])expect(()=>d.inspectBodyAnchors(q as string)).toThrow();expect(d.package.toBytes()).toEqual(before);
});

test('write and serialization failure roll back inserted text and keep the same anchor usable',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=await fixture(),a=d.inspectBodyAnchors()[0]!,before=d.package.toBytes(),pkg=(d as any).opcPackage as OpcPackage,original=pkg[stage].bind(pkg);if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected');};else pkg.toBytes=()=>{throw Error('injected');};try{expect(()=>d.insertParagraphAfter(a,'bad')).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}expect(d.package.toBytes()).toEqual(before);expect(d.insertParagraphAfter(a,'retry').text).toBe('retry');}
});

test('anchor insertion preserves aliased main XML UTF8 BOM and both UTF16 byte orders after reopen',async()=>{
 for(const codec of ['utf8','le','be']){const d=await fixture(s=>s.replaceAll('w:','q:').replaceAll('xmlns:w=','xmlns:q=')),pkg=await OpcPackage.open(d.package.toBytes()),raw=pkg.get(pkg.mainPart())!;if(codec==='utf8')pkg.set(pkg.mainPart(),new Uint8Array([239,187,191,...raw]));else{const text=pkg.text(pkg.mainPart()).replace('UTF-8','UTF-16'),b=new Uint8Array(2+text.length*2);b.set(codec==='le'?[255,254]:[254,255]);const v=new DataView(b.buffer);for(let i=0;i<text.length;i++)v.setUint16(2+2*i,text.charCodeAt(i),codec==='le');pkg.set(pkg.mainPart(),b);}const loaded=await Document.open(pkg.toBytes()),a=loaded.inspectBodyAnchors('Introduction')[0]!;loaded.insertParagraphAfter(a,'雪😀');const after=await Document.open(await loaded.save());expect(after.paragraphs.map(p=>p.text)).toEqual(['Introduction','雪😀','Customer context','Delivery approach','Use iterative delivery']);expect([...after.package.get('word/document.xml')!.slice(0,codec==='utf8'?3:2)]).toEqual(codec==='utf8'?[239,187,191]:codec==='le'?[255,254]:[254,255]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(after.package.get(n)).toEqual(pkg.get(n));}
});

test('discovery and insertion have separate output bounds with unchanged state on refusal',async()=>{
 const d=await fixture(),source=xml(d);d.package.setPart('word/document.xml',new TextEncoder().encode(source.replace('</w:body>',' '.repeat(8*1024*1024-source.length-1)+'</w:body>')));const loaded=await Document.open(d.package.toBytes()),a=loaded.inspectBodyAnchors()[0]!,before=loaded.package.toBytes();expect(()=>loaded.insertParagraphAfter(a,'too large')).toThrow();expect(loaded.package.toBytes()).toEqual(before);expect(loaded.inspectBodyAnchors()[0]!.text).toBe('Introduction');
});

test('three canonical body-anchor cases reject false headings empty filters and corrupted saved insertion',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings,cleanup}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/body-anchors.ts');const path='workflows/docx/anchor-discovery.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const run=async(active=bindings,text=source)=>{const feature=selectSharedScenarios(path,text,scenarioIds);try{return await executeAcceptance({root:'.',features:[feature],counts:{features:count(1),scenarios:count(3),cases:count(3),steps:count(16)}},active,'anchors-canonical');}finally{await cleanup();}};
 const good=await run();expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(3);expect(good.counts.cases.planned).toBe(2);
 for(const [step,mutate]of [
  ['Word anchors are listed without a query',(s:any)=>{s.anchors=s.anchors.map((a:any)=>({...a,type:'paragraph'}));}],
  ['Word anchors are listed with query "delivery"',(s:any)=>{s.anchors=[];}],
  ['Word anchors are listed with query "delivery"',(s:any)=>{s.anchors=s.document.inspectBodyAnchors();}],
  ['"Inserted after discovered anchor" is inserted after that anchor in the source document',async(s:any)=>{const d=await Document.open(s.path);d.paragraphs[3]!.setText('Wrong');await d.save(s.path);}],
 ] as const){const bad=bindings.map(b=>b.pattern.test(step)?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);await mutate(c.state);}}:b);const result=await run(bad);expect(result.counts.cases.failed).toBeGreaterThan(0);expect(result.counts.steps.failed).toBeGreaterThan(0);expect(result.counts.steps.undefined).toBe(0);expect(result.counts.steps.ambiguous).toBe(0);}
 const wrong=await run(bindings,source.replace('type "section_heading" and text "Introduction"','type "section_heading" and text "Wrong"'));expect(wrong.counts.cases.failed).toBe(1);expect(wrong.counts.steps.undefined).toBe(0);
});

test('saved anchor fixtures are cleaned after both successful cases and failed insertion reads',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings,cleanup}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/body-anchors.ts'),{lstat}=await import('node:fs/promises');const path='workflows/docx/anchor-discovery.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),roots:string[]=[],count=(n:number)=>({implemented:n,planned:0,total:n});
 const capture=bindings.map(b=>b.pattern.test('a saved Word document has headings "Introduction" and "Delivery approach" with paragraphs "Current intro" and "Current delivery"')||b.pattern.test('a saved Word document has headings "Introduction" and "Delivery approach" and paragraphs "Customer context paragraph" and "Use iterative delivery"')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);roots.push((c.state as {root:string}).root);}}:b);
 for(const text of [source,source.replace('shows "Inserted after discovered anchor" immediately','shows "Wrong" immediately')]){let result;try{result=await executeAcceptance({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(3),cases:count(3),steps:count(16)}},capture,'anchors-cleanup');}finally{await cleanup();}expect(result!.counts.cases.failed).toBe(text===source?0:1);}
 expect(roots).toHaveLength(6);for(const root of roots)await expect(lstat(root)).rejects.toMatchObject({code:'ENOENT'});
});
