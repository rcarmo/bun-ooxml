import {test,expect} from 'bun:test';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/style-authoring.ts';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';

test('paragraph style authoring executes24 saved outcomes and atomic refusals',async()=>{
 const path='workflows/docx/style-authoring.feature',f=parseFeature('references/fixtures-ooxml/'+path,(await Bun.file(join(fixturesRoot(),path)).text()).replace(/^@planned/m,'@implemented @bun')),rows=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(rows.length),steps:count(rows.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inv,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(24);
});

import {Document,W_NS as W} from '../../src/docx/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import {authoringDocument} from '../acceptance/style-authoring.ts';
import {STYLE_REL,STYLE_TYPE} from '../acceptance/paragraph-style.ts';

test('definition creation preserves paragraph, table-cell and span handles until a text edit',async()=>{
 const d=Document.create();const table=d.addTable(1,1);table.cell(0,0).text='cell';const p=d.paragraphs.find(p=>p.text==='cell')!,span=p.find('cell')[0]!,cell=table.cell(0,0),before=d.package.get('word/document.xml'),version=d.currentVersion();
 d.addParagraphStyle('Custom',{name:'Custom',bold:true});expect(d.currentVersion()).toBe(version);expect(p.text).toBe('cell');expect(cell.text).toBe('cell');expect(d.package.get('word/document.xml')).toEqual(before);await span.replace('updated');expect(table.cell(0,0).text).toBe('updated');const added=d.addParagraph('new',{style:'Custom'});expect(added.styleId).toBe('Custom');const q=await Document.open(await d.save());expect(q.paragraphs.at(-1)!.styleId).toBe('Custom');
});
test('styles-only update retains UTF16 LE and BE declarations, BOM and old definitions',async()=>{
 for(const little of [true,false]){
 const d=await authoringDocument('existing'),p=await OpcPackage.open(await d.save()),xml=p.text('word/styles.xml');const text='<?xml version="1.0" encoding="UTF-16"?>'+xml;const bytes=new Uint8Array(2+text.length*2),view=new DataView(bytes.buffer);view.setUint16(0,0xfeff,little);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),little);p.set('word/styles.xml',bytes);const q=await Document.open(p.toBytes());q.addParagraphStyle('Custom',{name:'π style',italic:false});const saved=await OpcPackage.open(await q.save());expect([...saved.get('word/styles.xml')!.slice(0,2)]).toEqual(little?[255,254]:[254,255]);expect(saved.text('word/styles.xml')).toContain('encoding="UTF-16"');for(const n of parseXml(xml).root.children)expect(saved.text('word/styles.xml')).toContain(xml.slice(n.start,n.end));expect(saved.get('word/document.xml')).toEqual(p.get('word/document.xml'));
 }
});
test('final serialization failure rolls back new registry graph and existing registry append',async()=>{
 for(const kind of ['absent','existing']){
  const d=await authoringDocument(kind);d.addParagraph('earlier');const before=await d.save(),handle=d.paragraphs[0],version=d.currentVersion(),serialize=OpcPackage.prototype.toBytes;let count=0;
  OpcPackage.prototype.toBytes=function(){if(++count===(kind==='absent'?3:1))throw new Error('injected final style serialization');return serialize.call(this);};
  try{expect(()=>d.addParagraphStyle('New',{name:'New'})).toThrow('injected final style serialization');}finally{OpcPackage.prototype.toBytes=serialize;}
  expect(count).toBe(kind==='absent'?3:1);expect(await d.save()).toEqual(before);expect(d.currentVersion()).toBe(version);expect(d.paragraphs[0]).toBe(handle);d.addParagraphStyle('New',{name:'New'});expect(d.paragraphs.at(-1)!.text).toBe('earlier');
 }
});
test('input data rejects unknown keys, accessors and invalid XML without package changes',async()=>{
 const d=await authoringDocument('existing'),before=await d.save();let getters=0;const accessor={name:'ok',get bold(){getters++;return true;}};
 for(const [id,opts]of [['', {name:'x'}],[' x',{name:'x'}],['x',{name:' '}],['x',{name:'bad\u0001'}],['bad\u0001',{name:'x'}],['x',{name:'x',bold:null}],['x',{name:'x',basedOn:1}],['x',{name:'x',basedOn:''}],['x',{name:'x',default:true}],['x',accessor],['x',Object.assign(Object.create({basedOn:'Heading1'}),{name:'x'})],['x',{name:'x',[Symbol('bad')]:true}]] as const){expect(()=>d.addParagraphStyle(id,opts as any)).toThrow();expect(await d.save()).toEqual(before);}
 expect(getters).toBe(0);
});
test('all registry IDs must be unique for authoring, including unselected duplicates',async()=>{
 const d=await authoringDocument('existing'),p=await OpcPackage.open(await d.save());p.set('word/styles.xml',p.text('word/styles.xml').replace('</w:styles>','<w:style w:type="character" w:styleId="Body"/></w:styles>'));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.addParagraphStyle('New',{name:'New'})).toThrow('Duplicate registry style ID');expect(await q.save()).toEqual(before);
});
test('selected base chains require paragraph targets and refuse multiple or foreign links',async()=>{
 for(const next of ['<w:basedOn w:val="Body"/><w:basedOn w:val="Body"/>','<x:basedOn xmlns:x="urn:wrong" x:val="Body"/>','<w:basedOn w:val="Body"><w:other/></w:basedOn>','<w:basedOn w:val="Emphasis"/>']){
  const d=await authoringDocument('existing'),p=await OpcPackage.open(await d.save());p.set('word/styles.xml',p.text('word/styles.xml').replace('<w:name w:val="Heading One"/>','<w:name w:val="Heading One"/>'+next));const q=await Document.open(p.toBytes()),before=await q.save();expect(()=>q.addParagraphStyle('New',{name:'New',basedOn:'Heading1'})).toThrow();expect(await q.save()).toEqual(before);
 }
 const d=await authoringDocument('existing');d.addParagraphStyle('One',{name:'One',basedOn:'Heading1'});d.addParagraphStyle('Two',{name:'Two',basedOn:'One'});d.paragraphs[0]!.setStyle('Two');expect((await Document.open(await d.save())).paragraphs[0]!.styleId).toBe('Two');
});
test('new styles honor arbitrary relationship target paths and preserve registry metadata',async()=>{
 const d=Document.create();d.addParagraph('x');const p=await OpcPackage.open(await d.save());const xml=`<s:styles xmlns:s="${W}" data="keep"><s:docDefaults/><s:latentStyles s:count="0"/><s:style s:type="paragraph" s:styleId="Body"><s:name s:val="Body"/></s:style></s:styles>`;
 addPart(p,'custom/styles.xml',xml,STYLE_TYPE);addRelationship(p,p.mainPart(),STYLE_REL,'../custom/styles.xml');const q=await Document.open(p.toBytes());expect(q.addParagraphStyle('New',{name:'New',basedOn:'Body'})).toEqual({styleId:'New',partName:'custom/styles.xml'});const saved=await OpcPackage.open(await q.save());expect(saved.text('custom/styles.xml')).toContain('<s:docDefaults/><s:latentStyles s:count="0"/>');expect(saved.text('custom/styles.xml')).toContain('data="keep"');expect(saved.get('word/_rels/document.xml.rels')).toEqual(p.get('word/_rels/document.xml.rels'));expect(saved.get('[Content_Types].xml')).toEqual(p.get('[Content_Types].xml'));
});
test('XML attribute escapes and dollar replacement tokens are literal in self-closing registries',async()=>{
 const d=await authoringDocument('empty'),id='A&"\'$&',name='Quote " & π $` $\' $&';d.addParagraphStyle(id,{name,bold:false,italic:true});const p=await OpcPackage.open(await d.save()),def=elements(parseXml(p.text('word/styles.xml')),'style',W)[0]!;expect(attribute(def,'styleId',W)).toBe(id);expect(attribute(elements(def,'name',W)[0]!,'val',W)).toBe(name);d.paragraphs[0]!.setStyle(id);expect((await Document.open(await d.save())).paragraphs[0]!.styleId).toBe(id);
});
test('disabled protection allows authoring but external or later enforcing settings refuse',async()=>{
 for(const kind of ['external','disabled','later']){
  const d=await authoringDocument('existing'),p=await OpcPackage.open(await d.save());const type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
  if(kind==='external')addRelationship(p,p.mainPart(),type,'https://example.invalid/settings.xml',{external:true});else{addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="0"/>${kind==='later'?'<w:documentProtection w:enforcement="1"/>':''}</w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),type,'settings.xml');}
  const q=await Document.open(p.toBytes()),before=await q.save();if(kind==='disabled')expect(q.addParagraphStyle('New',{name:'New'}).styleId).toBe('New');else{expect(()=>q.addParagraphStyle('New',{name:'New'})).toThrow();expect(await q.save()).toEqual(before);}
 }
});
test('disk save reopen retains authored base and flags without dirtying document text',async()=>{
 const {mkdtemp,rm}=await import('node:fs/promises'),{tmpdir}=await import('node:os'),{join}=await import('node:path'),root=await mkdtemp(join(tmpdir(),'bun-style-authoring-'));try{const d=await authoringDocument('existing');d.addParagraphStyle('New',{name:'New',basedOn:'Heading1',bold:true,italic:false});d.paragraphs[0]!.setStyle('New');const path=join(root,'authored.docx');await d.save(path);const q=await Document.open(path);expect(q.paragraphs[0]!.styleId).toBe('New');expect(q.paragraphs[0]!.text).toBe('Styled π text');const p=await OpcPackage.open(await q.save()),def=elements(parseXml(p.text('word/styles.xml')),'style',W).find(n=>attribute(n,'styleId',W)==='New')!;expect(attribute(elements(def,'basedOn',W)[0]!,'val',W)).toBe('Heading1');expect(attribute(elements(def,'b',W)[0]!,'val',W)).toBe('1');expect(attribute(elements(def,'i',W)[0]!,'val',W)).toBe('0');}finally{await rm(root,{recursive:true,force:true});}
});
