import {test,expect} from 'bun:test';
import {Document,OpcPackage} from '../../src/index.ts';
import {mkdtemp,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
async function styled(id:string|null){const d=Document.create();if(id!==null)d.addParagraphStyle(id,{name:'Unrelated display name'});d.addParagraph('Heading text');d.paragraphs[0]!.setStyle(id);return d;}

test('direct Heading1 through Heading9 classify by identifier without changing package bytes',async()=>{
 for(let level=1;level<=9;level++){const d=await styled('Heading'+level),p=d.paragraphs[0]!,before=d.package.toBytes();expect(p.styleId).toBe('Heading'+level);expect(p.isHeading).toBe(true);expect(p.headingLevel).toBe(level);expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Heading text');}
});

test('absent ordinary and misleading style IDs do not infer a heading',async()=>{
 for(const id of [null,'Normal','Title','heading1','HEADING1','Heading0','Heading10','Heading01','Heading1Suffix','Heading 1','CustomHeading']){const d=await styled(id),before=d.package.toBytes(),p=d.paragraphs[0]!;expect(p.styleId).toBe(id??undefined);expect(p.isHeading).toBe(false);expect(p.headingLevel).toBe(0);expect(d.package.toBytes()).toEqual(before);}
});

test('direct style classification does not compute base-style outline or body-anchor semantics',()=>{
 const d=Document.create();d.addParagraphStyle('Heading1',{name:'First heading'});d.addParagraphStyle('Custom',{name:'Heading 1',basedOn:'Heading1'});d.addParagraph('Custom paragraph',{style:'Custom'}).setProperties({outlineLevel:0});
 let p=d.paragraphs[0]!;expect(p.headingLevel).toBe(0);expect(p.isHeading).toBe(false);expect(d.inspectBodyAnchors('Custom')[0]!.type).toBe('section_heading');
 p.setStyle('Heading1');p=d.paragraphs[0]!;p.setProperties({outlineLevel:9});p=d.paragraphs[0]!;expect(p.headingLevel).toBe(1);expect(p.isHeading).toBe(true);expect(d.inspectBodyAnchors('Custom')[0]!.type).toBe('paragraph');
});

test('style classification survives path save reopen without changing text or unrelated members',async()=>{
 const d=await styled('Heading9'),before=d.package.parts,dir=await mkdtemp(join(tmpdir(),'heading-classification-'));try{const path=join(dir,'heading.docx');await d.save(path);const read=await Document.open(path);expect(read.paragraphs[0]!.headingLevel).toBe(9);expect(read.paragraphs[0]!.isHeading).toBe(true);expect(read.paragraphs[0]!.styleId).toBe('Heading9');expect(read.paragraphs[0]!.text).toBe('Heading text');for(const[n,b]of before)expect(read.package.get(n)).toEqual(b);}finally{await rm(dir,{recursive:true,force:true});}
});

test('changed or externally modified paragraph handles refuse classification instead of serving cached headings',async()=>{
 const d=await styled('Heading1'),held=d.paragraphs[0]!;held.setStyle(null);expect(()=>held.headingLevel).toThrow(expect.objectContaining({code:'docx-stale-paragraph'}));expect(()=>held.isHeading).toThrow(expect.objectContaining({code:'docx-stale-paragraph'}));expect(d.paragraphs[0]!.headingLevel).toBe(0);
 const fresh=d.paragraphs[0]!,before=new TextDecoder().decode(d.package.get('word/document.xml'));d.package.setPart('word/document.xml',new TextEncoder().encode(before.replace('Heading text','External text')));const bytes=d.package.toBytes();expect(()=>fresh.isHeading).toThrow();expect(()=>fresh.headingLevel).toThrow();expect(d.package.toBytes()).toEqual(bytes);
});

test('classification reads literal imported IDs without requiring style definitions but refuses ambiguous direct metadata',async()=>{
 const d=Document.create();d.addParagraph('Imported');const base=await OpcPackage.open(d.package.toBytes());
 for(const metadata of ['<w:pStyle w:val="Heading3"/>','<w:pStyle w:val="Heading1"/><w:pStyle w:val="Heading2"/>','<w:pStyle val="Heading1"/>']){const p=await OpcPackage.open(base.toBytes());p.set(p.mainPart(),p.text(p.mainPart()).replace(/<w:p(?:\s[^>]*)?>/,open=>open+'<w:pPr>'+metadata+'</w:pPr>'));const loaded=await Document.open(p.toBytes()),before=loaded.package.toBytes();if(metadata.includes('Heading3')){expect(loaded.paragraphs[0]!.headingLevel).toBe(3);expect(loaded.paragraphs[0]!.isHeading).toBe(true);}else{expect(()=>loaded.paragraphs[0]!.headingLevel).toThrow();expect(()=>loaded.paragraphs[0]!.isHeading).toThrow();}expect(loaded.package.toBytes()).toEqual(before);}
});

test('aliased UTF16 direct style reads without editing encoding or root namespace bindings',async()=>{
 const d=await styled('Heading2'),p=await OpcPackage.open(d.package.toBytes());const text=p.text(p.mainPart()).replaceAll('w:','q:').replace('xmlns:w=','xmlns:q=').replace('UTF-8','UTF-16');const bytes=new Uint8Array(2+text.length*2);bytes.set([255,254]);const v=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)v.setUint16(2+i*2,text.charCodeAt(i),true);p.set(p.mainPart(),bytes);const loaded=await Document.open(p.toBytes()),before=loaded.package.toBytes();expect(loaded.paragraphs[0]!.headingLevel).toBe(2);expect(loaded.paragraphs[0]!.isHeading).toBe(true);expect(loaded.package.toBytes()).toEqual(before);
});

test('empty style retains strict Bun policy and all five shared getter-profile cases stay planned',async()=>{
 const d=await styled('Heading1'),held=d.paragraphs[0]!,before=d.package.toBytes();expect(()=>held.setStyle('')).toThrow(expect.objectContaining({code:'docx-style-argument'}));expect(d.package.toBytes()).toEqual(before);expect(held.headingLevel).toBe(1);held.setStyle(null);const p=d.paragraphs[0]!;expect(p.styleId).toBeUndefined();expect(p.isHeading).toBe(false);expect(p.headingLevel).toBe(0);
 const {inventoryFeatures}=await import('../../scripts/gherkin.ts'),inv=await inventoryFeatures(process.cwd()),profile=inv.features.flatMap(f=>f.scenarios).find(s=>s.scenarioId==='@id-docx-go-paragraph-style-getters')!;expect(profile.lifecycle).toBe('planned');expect(profile.cases).toHaveLength(5);expect(inv.counts.cases.implemented).toBe(551);expect(inv.counts.cases.planned).toBe(46);
});
