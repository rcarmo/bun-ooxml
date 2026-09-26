import {test,expect} from 'bun:test';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/text-box.ts';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';

test('positioned text boxes execute23 saved outcomes and atomic refusals',async()=>{
 const path='workflows/pptx/text-box.feature',f=parseFeature('references/fixtures-ooxml/'+path,(await Bun.file(join(fixturesRoot(),path)).text()).replace(/^@planned/m,'@implemented @bun')),rows=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(rows.length),steps:count(rows.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inv,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(23);
});

import {Presentation} from '../../src/pptx/index.ts';
import {OpcPackage} from '../../src/opc/index.ts';
import {boxDeck,geometry,P,A} from '../acceptance/text-box.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
import {fixturePath,F} from '../../scripts/fixture-inputs.ts';

test('text-box creation stales same-slide anchors and tables but not other slides',async()=>{
 const d=await boxDeck('plain'),s=d.slides[0]!,other=d.slides[1]!;s.addTable(1,1,geometry);const table=s.tables[0]!,cell=table.cell(0,0),anchor=s.inspectText('old')[0]!.anchor,otherAnchor=other.inspectText('other')[0]!.anchor;
 s.addTextBox('new text',geometry);const before=d.package.toBytes();expect(()=>s.replaceTextAt(anchor,'Existing','Old')).toThrow('stale');expect(()=>cell.text).toThrow('stale');expect(()=>table.rows).toThrow('stale');expect(d.package.toBytes()).toEqual(before);expect(s.tables[0]!.rows).toBe(1);other.replaceTextAt(otherAnchor,'Other','Changed');const box=s.inspectText('fresh').find(p=>p.text==='new text')!;s.replaceTextAt(box.anchor,'new','updated');const q=await Presentation.open(d.package.toBytes());expect(q.slides[0]!.inspectText('reopened').some(p=>p.text==='updated text')).toBe(true);expect(q.slides[1]!.inspectText('reopened')[0]!.text).toBe('Changed slide');
});
test('serialization failure restores prior edits, versions and still-usable anchors',async()=>{
 const d=await boxDeck('plain'),s=d.slides[0]!;s.addTextBox('earlier',geometry);const before=d.package.toBytes(),version=d.currentSlideVersion(s.partName),anchor=s.inspectText('before')[0]!.anchor,original=OpcPackage.prototype.toBytes;OpcPackage.prototype.toBytes=function(){throw Error('injected box serialization');};try{expect(()=>s.addTextBox('later',geometry)).toThrow('injected box serialization');}finally{OpcPackage.prototype.toBytes=original;}
 expect(d.package.toBytes()).toEqual(before);expect(d.currentSlideVersion(s.partName)).toBe(version);s.replaceTextAt(anchor,'Existing','Kept');expect(s.inspectText('after')[0]!.text).toBe('Kept title');
});
test('slide text-box writing preserves UTF16 LE and BE declarations and BOM',async()=>{
 for(const little of [true,false]){const d=await boxDeck('plain'),s=d.slides[0]!,xml=d.package.text(s.partName).replace('encoding="UTF-8"','encoding="UTF-16"'),bytes=new Uint8Array(2+xml.length*2),view=new DataView(bytes.buffer);view.setUint16(0,0xfeff,little);for(let i=0;i<xml.length;i++)view.setUint16(2+2*i,xml.charCodeAt(i),little);d.package.set(s.partName,bytes);s.addTextBox('π text',geometry);const saved=await Presentation.open(d.package.toBytes());expect([...saved.package.get(s.partName)!.slice(0,2)]).toEqual(little?[255,254]:[254,255]);expect(saved.package.text(s.partName)).toContain('encoding="UTF-16"');expect(saved.slides[0]!.inspectText('read').at(-1)!.text).toBe('π text');}
});
test('new shapes bind namespaces locally even when p/a are absent or shadowed',async()=>{
 const d=await boxDeck('alias'),s=d.slides[0]!;d.package.set(s.partName,d.package.text(s.partName).replace('<q:spTree>','<q:spTree xmlns:p="urn:wrong-p" xmlns:a="urn:wrong-a">'));s.addTextBox('literal $& <&> "π"\t ',geometry,{name:'literal $& "π"',bold:false,italic:true});const root=parseXml(d.package.text(s.partName)),sp=elements(root,'sp',P).at(-1)!;expect(elements(sp,'cNvPr',P)[0]!.attributes.name).toBe('literal $& "π"');expect(elements(sp,'t',A)[0]!.text).toBe('literal $& <&> "π"\t ');expect(elements(sp,'rPr',A)[0]!.attributes).toEqual({b:'0',i:'1'});
});
test('geometry and options reject accessors, inherited values and unknown fields without invoking getters',async()=>{
 const d=await boxDeck('plain'),s=d.slides[0]!,before=d.package.toBytes();let calls=0;
 for(const [g,o]of [[{...geometry,get x(){calls++;return 1;}},{}],[geometry,{get bold(){calls++;return true;}}],[Object.assign(Object.create({x:1}),{y:1,width:1,height:1}),{}],[{...geometry,rotation:1},{}],[geometry,{name:''}],[geometry,{fontSize:12}],[geometry,{italic:null}],[{...geometry,width:NaN},{}],[{...geometry,height:Infinity},{}],[geometry,{[Symbol('extra')]:1}]] as const){expect(()=>s.addTextBox('x',g as any,o as any)).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(calls).toBe(0);
});
test('shape-ID allocation is lexical, slide-local and counts nested existing shapes',async()=>{
 const d=await boxDeck('nested-id'),s=d.slides[0]!;expect(s.addTextBox('one',geometry).shapeId).toBe(96);expect(s.addTextBox('two',geometry).shapeId).toBe(97);expect(d.slides[1]!.addTextBox('other',geometry).shapeId).toBe(3);
 for(const id of ['2.5','2e2','+2','-2','0','2147483648','0001']){const q=await boxDeck('plain'),slide=q.slides[0]!;q.package.set(slide.partName,q.package.text(slide.partName).replace('id="2"',`id="${id}"`));const before=q.package.toBytes();expect(()=>slide.addTextBox('x',geometry)).toThrow();expect(q.package.toBytes()).toEqual(before);}
});
test('tree ambiguity, lexical barriers and malformed shape identity refuse atomically',async()=>{
 for(const [from,to]of [['<p:cSld>','<p:cSld/><p:cSld>'],['<p:spTree>','<p:spTree><!--barrier-->'],['<p:nvSpPr>','<p:nvSpPr><p:cNvPr id="8"/>'],['<p:cNvPr id="2" name="Title 1"/>','']] as const){const d=await boxDeck('plain'),s=d.slides[0]!,xml=d.package.text(s.partName);expect(xml.includes(from)).toBe(true);d.package.set(s.partName,xml.replace(from,to));const before=d.package.toBytes();expect(()=>s.addTextBox('x',geometry)).toThrow();expect(d.package.toBytes()).toEqual(before);}
});
test('external slide reorder invalidates old slide identity while raw same-slide edits are read live',async()=>{
 const d=await boxDeck('plain'),s=d.slides[0]!;d.package.set(s.partName,d.package.text(s.partName).replace('Existing title','External text'));s.addTextBox('new',geometry);expect(s.inspectText('live')[0]!.text).toBe('External text');const main=d.package.mainPart(),xml=d.package.text(main),doc=parseXml(xml),ids=elements(doc,'sldId',P);d.package.set(main,xml.replace(xml.slice(ids[0]!.start,ids[1]!.end),xml.slice(ids[1]!.start,ids[1]!.end)+xml.slice(ids[0]!.start,ids[0]!.end)));const before=d.package.toBytes();expect(()=>s.addTextBox('refused',geometry)).toThrow('Slide order');expect(d.package.toBytes()).toEqual(before);
});
test('real presentation adds a text box and saves/reopens without changing other payloads',async()=>{
 const d=await Presentation.open(await fixturePath(F.goSlides.notes)),s=d.slides[0]!,before=new Map(d.package.names().map(n=>[n,d.package.get(n)!]));const receipt=s.addTextBox('Native textbox',geometry,{bold:true});const q=await Presentation.open(d.package.toBytes());expect(q.slides[0]!.inspectText('read').some(p=>p.text==='Native textbox')).toBe(true);expect(receipt.shapeId).toBeGreaterThan(1);for(const [n,bytes]of before)if(n!==s.partName)expect(q.package.get(n)).toEqual(bytes);
});
test('disk save retains multiline paragraphs, direct flags and extension tail order',async()=>{
 const {mkdtemp,rm}=await import('node:fs/promises'),{tmpdir}=await import('node:os'),{join}=await import('node:path'),root=await mkdtemp(join(tmpdir(),'bun-text-box-'));try{const d=await boxDeck('extension-tail'),s=d.slides[0]!;s.addTextBox('one\n\ntwo',geometry,{bold:true,italic:false});const path=join(root,'boxes.pptx');await d.save(path);const q=await Presentation.open(path),tree=elements(parseXml(q.package.text(s.partName)),'spTree',P)[0]!;expect(tree.children.at(-1)!.localName).toBe('extLst');expect(tree.children.at(-2)!.localName).toBe('sp');expect(q.slides[0]!.inspectText('read').slice(1).map(p=>p.text)).toEqual(['one','','two']);expect(q.slides[0]!.inspectText('read').slice(1).map(p=>p.runs[0]!.attrs)).toEqual([{b:'1',i:'0'},{b:'1',i:'0'},{b:'1',i:'0'}]);}finally{await rm(root,{recursive:true,force:true});}
});
