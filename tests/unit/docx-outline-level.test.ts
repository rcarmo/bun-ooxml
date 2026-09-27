import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';import {addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(change=(xml:string)=>xml){const d=Document.create();d.addParagraph('Outline text');const p=await OpcPackage.open(d.package.toBytes());const source=p.text(p.mainPart()).replace(`<w:p xmlns:w="${W}">`,'<w:p>');p.set(p.mainPart(),change(source));return Document.open(p.toBytes());}
const xml=(d:Document)=>new TextDecoder().decode(d.package.get('word/document.xml'));

test('all ten direct outline levels survive independent path roundtrips without changing text or other parts',async()=>{
 const root=await mkdtemp(join(tmpdir(),'outline-level-'));try{for(let level=0;level<=9;level++){const d=await fixture(),p=d.paragraphs[0]!,before=d.package.parts;expect(p.directProperties().outlineLevel).toBeNull();expect(p.setProperties({outlineLevel:level})).toEqual({changed:1});expect(()=>p.directProperties()).toThrow(expect.objectContaining({code:'docx-stale-paragraph'}));const path=join(root,`level-${level}.docx`);await d.save(path);const after=await Document.open(path);expect(after.paragraphs[0]!.directProperties().outlineLevel).toBe(level);expect(after.paragraphs[0]!.text).toBe('Outline text');expect(elements(parseXml(xml(after)),'outlineLvl',W)[0]!.attributes['w:val']).toBe(String(level));for(const[n,b]of before)if(n!=='word/document.xml')expect(after.package.get(n)).toEqual(b);}}finally{await rm(root,{recursive:true,force:true});}
});

test('absent null explicit nine and zero stay distinct and lexical no-ops preserve handles',async()=>{
 const d=await fixture(x=>x.replace('<w:p>','<w:p><w:pPr><w:outlineLvl w:val = \'00\'/></w:pPr>')),p=d.paragraphs[0]!,before=d.package.toBytes();expect(p.directProperties().outlineLevel).toBe(0);expect(p.setProperties({outlineLevel:0})).toEqual({changed:0});expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Outline text');expect(p.setProperties({outlineLevel:null})).toEqual({changed:1});expect(d.paragraphs[0]!.directProperties().outlineLevel).toBeNull();expect(xml(d)).not.toContain('outlineLvl');d.paragraphs[0]!.setProperties({outlineLevel:9});expect(d.paragraphs[0]!.directProperties().outlineLevel).toBe(9);const now=d.package.toBytes(),current=d.paragraphs[0]!;expect(current.setProperties({outlineLevel:9})).toEqual({changed:0});expect(d.package.toBytes()).toEqual(now);expect(current.text).toBe('Outline text');
});

test('outline insertion respects property order and keeps direct styles and unrelated XML fragments',async()=>{
 const d=await fixture(x=>x.replace('<w:p>','<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="right"/><w:divId w:val="0"/></w:pPr>')),before=xml(d);expect(d.paragraphs[0]!.setProperties({outlineLevel:2})).toEqual({changed:1});const after=xml(d),pr=elements(parseXml(after),'pPr',W)[0]!;expect(pr.children.map(n=>n.localName)).toEqual(['pStyle','jc','outlineLvl','divId']);expect(after.replace(/<w:outlineLvl[^>]*\/>/,'')).toBe(before);expect(d.paragraphs[0]!.styleId).toBe('Heading1');expect(d.paragraphs[0]!.directProperties()).toMatchObject({alignment:'right',outlineLevel:2});
});

test('missing and self-closing paragraph properties allow direct outline edits without style synthesis',async()=>{
 for(const pr of ['', '<w:pPr/>']){const d=await fixture(x=>x.replace('<w:p>','<w:p>'+pr));d.paragraphs[0]!.setProperties({outlineLevel:0});const after=await Document.open(await d.save());expect(after.paragraphs[0]!.directProperties().outlineLevel).toBe(0);expect(after.paragraphs[0]!.styleId).toBeUndefined();expect(after.package.has('word/styles.xml')).toBe(false);expect(after.paragraphs[0]!.text).toBe('Outline text');}
 const d=await fixture(x=>x.replace('<w:p><w:r><w:t>Outline text</w:t></w:r></w:p>','<w:p/>'));expect(d.paragraphs[0]!.text).toBe('');d.paragraphs[0]!.setProperties({outlineLevel:1});expect((await Document.open(await d.save())).paragraphs[0]!.directProperties().outlineLevel).toBe(1);
});

test('invalid outline values and a late invalid field refuse the complete patch with bytes and handles intact',async()=>{
 const d=await fixture(),p=d.paragraphs[0]!,before=d.package.toBytes();for(const value of [-1,10,0.5,NaN,Infinity,'0',true,{},new Number(1)]){expect(()=>p.setProperties({alignment:'center',outlineLevel:value as number})).toThrow(expect.objectContaining({code:'docx-paragraph-properties-unsupported'}));expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Outline text');expect(p.directProperties().alignment).toBeNull();}
 let called=false;const patch=Object.defineProperty({},'outlineLevel',{enumerable:true,get(){called=true;return 1;}});expect(()=>p.setProperties(patch)).toThrow();expect(called).toBe(false);expect(d.package.toBytes()).toEqual(before);
});

test('duplicate misordered malformed lexical or wrongly qualified outline metadata refuses reads writes and removal',async()=>{
 for(const pr of [
 '<w:outlineLvl w:val="0"/><w:outlineLvl w:val="1"/>','<w:outlineLvl/>','<w:outlineLvl val="1"/>','<w:outlineLvl w:val="10"/>','<w:outlineLvl w:val="-1"/>','<w:outlineLvl w:val="x"/>','<w:outlineLvl w:val="0" w:extra="x"/>','<w:outlineLvl w:val="0"><!--keep--></w:outlineLvl>','<w:outlineLvl w:val="0"><w:bad/></w:outlineLvl>','<w:outlineLvl w:val="0"/><w:jc w:val="left"/>','<w:outlineLvl w:val="0"/><w:pPrChange/>',
 ]){const d=await fixture(x=>x.replace('<w:p>','<w:p><w:pPr>'+pr+'</w:pPr>')),p=d.paragraphs[0]!,before=d.package.toBytes();expect(()=>p.directProperties()).toThrow();expect(()=>p.setProperties({outlineLevel:0})).toThrow();expect(()=>p.setProperties({outlineLevel:null})).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('protected no-op external main changes and field-bearing paragraphs refuse outline mutation',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const locked=await Document.open(pkg.toBytes()),before=locked.package.toBytes();expect(locked.paragraphs[0]!.directProperties().outlineLevel).toBeNull();expect(()=>locked.paragraphs[0]!.setProperties({outlineLevel:null})).toThrow();expect(locked.package.toBytes()).toEqual(before);
 const p=d.paragraphs[0]!;d.package.setPart('word/document.xml',new TextEncoder().encode(xml(d)+' '));const changed=d.package.toBytes();expect(()=>p.setProperties({outlineLevel:0})).toThrow();expect(d.package.toBytes()).toEqual(changed);
 const field=await fixture(x=>x.replace('<w:t>Outline text</w:t>','<w:fldChar w:fldCharType="begin"/>')),fb=field.package.toBytes();expect(()=>field.paragraphs[0]!.setProperties({outlineLevel:0})).toThrow();expect(field.package.toBytes()).toEqual(fb);
});

test('outline edits retain table grid and sibling text but expire paragraph cell and span snapshots',async()=>{
 const d=Document.create();d.addTable(1,2);d.tables[0]!.cell(0,0).text='Left';d.tables[0]!.cell(0,1).text='Right';const t=d.tables[0]!,cell=t.cell(0,0),p=d.paragraphs[0]!,span=p.find('Left')[0]!;p.setProperties({outlineLevel:3});expect(()=>p.text).toThrow();expect(()=>cell.text).toThrow();expect(()=>span.replace('stale')).toThrow();expect([t.rows,t.columns,t.cell(0,1).text]).toEqual([1,2,'Right']);const after=await Document.open(await d.save());expect(after.paragraphs[0]!.directProperties().outlineLevel).toBe(3);expect(after.tables[0]!.cell(0,1).text).toBe('Right');
});

test('set and serialization faults roll back outline edits with live paragraph and cell handles',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=Document.create();d.addTable(1,1);d.tables[0]!.cell(0,0).text='Held';const p=d.paragraphs[0]!,cell=d.tables[0]!.cell(0,0),before=d.package.toBytes(),pkg=(d as any).opcPackage as OpcPackage,original=pkg[stage].bind(pkg);if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected');};else pkg.toBytes=()=>{throw Error('injected');};try{expect(()=>p.setProperties({outlineLevel:0})).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Held');expect(cell.text).toBe('Held');expect(p.setProperties({outlineLevel:0})).toEqual({changed:1});}
});

test('namespace aliases UTF8 BOM and UTF16 codecs survive outline edits and byte reopen',async()=>{
 for(const codec of ['utf8','le','be']){const d=await fixture(x=>x.replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('<q:p>','<q:p xmlns:w="urn:foreign">')),pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()),raw=pkg.get(pkg.mainPart())!;if(codec==='utf8')pkg.set(pkg.mainPart(),new Uint8Array([239,187,191,...raw]));else{const source=text.replace('UTF-8','UTF-16'),b=new Uint8Array(2+source.length*2);b.set(codec==='le'?[255,254]:[254,255]);const view=new DataView(b.buffer);for(let i=0;i<source.length;i++)view.setUint16(2+i*2,source.charCodeAt(i),codec==='le');pkg.set(pkg.mainPart(),b);}const loaded=await Document.open(pkg.toBytes());loaded.paragraphs[0]!.setProperties({outlineLevel:8});const after=await Document.open(await loaded.save());expect(after.paragraphs[0]!.directProperties().outlineLevel).toBe(8);expect(after.paragraphs[0]!.text).toBe('Outline text');expect([...after.package.get('word/document.xml')!.slice(0,codec==='utf8'?3:2)]).toEqual(codec==='utf8'?[239,187,191]:codec==='le'?[255,254]:[254,255]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(after.package.get(n)).toEqual(pkg.get(n));}
});

test('complete output bounds refuse new outline metadata without invalidating the held paragraph',async()=>{
 const d=await fixture(),text=xml(d),next=text.replace('</w:body>',' '.repeat(8*1024*1024-text.length-1)+'</w:body>');d.package.setPart('word/document.xml',new TextEncoder().encode(next));const loaded=await Document.open(d.package.toBytes()),p=loaded.paragraphs[0]!,before=loaded.package.toBytes();expect(()=>p.setProperties({outlineLevel:0})).toThrow(expect.objectContaining({code:'XML_EDIT_UNSAFE'}));expect(loaded.package.toBytes()).toEqual(before);expect(p.directProperties().outlineLevel).toBeNull();expect(p.text).toBe('Outline text');
});
