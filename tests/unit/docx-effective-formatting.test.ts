import {test,expect} from 'bun:test';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/effective-formatting.ts';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
test('effective run formatting executes 25 saved outcomes and read-only refusals',async()=>{
 const path='workflows/docx/effective-formatting.feature',f=parseFeature('references/fixtures-ooxml/'+path,(await Bun.file(join(fixturesRoot(),path)).text()).replace(/^@planned/m,'@implemented @bun')),rows=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(rows.length),steps:count(rows.reduce((n,c)=>n+c.steps.length,0))}};
 const r=await executeAcceptance(inv,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(25);
});

import {Document} from '../../src/docx/index.ts';
import {OpcPackage} from '../../src/opc/index.ts';
import {formattingDocument,W} from '../acceptance/effective-formatting.ts';
const encoder=new TextEncoder();
async function changed(kind:string,part:string,update:(xml:string)=>string){const d=await formattingDocument(kind),p=await OpcPackage.open(d.package.toBytes());p.set(part,update(p.text(part)));return Document.open(p.toBytes());}
test('defaults combine with style toggles and provenance records every transition',async()=>{
 const d=await changed('toggle-chain','word/styles.xml',s=>s.replace('<w:style ','<w:docDefaults><w:rPrDefault><w:rPr><w:b/><w:i/></w:rPr></w:rPrDefault></w:docDefaults><w:style '));
 const row=d.paragraphs[0]!.effectiveRunFormatting()[0]!;
 expect([row.bold.value,row.italic.value]).toEqual([true,false]);
 expect(row.bold.contributions).toEqual([{source:'implicit',operation:'set',value:false,result:false},{source:'document-default',partName:'word/styles.xml',operation:'set',value:true,result:true},{source:'paragraph-style',partName:'word/styles.xml',styleId:'Base',operation:'toggle',value:true,result:false},{source:'paragraph-style',partName:'word/styles.xml',styleId:'Child',operation:'toggle',value:true,result:true}]);
});
test('an explicit style chain does not also apply the default paragraph style',async()=>{
 const d=await changed('inherited','word/styles.xml',s=>s.replace('</w:styles>','<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:i/></w:rPr></w:style></w:styles>'));
 const r=d.paragraphs[0]!.effectiveRunFormatting()[0]!;expect(r.paragraphStyleChain).toEqual(['Base','Child']);expect([r.bold.value,r.italic.value]).toEqual([true,true]);
});
test('inspection returns detached results and never invalidates a captured text span',async()=>{
 const d=await formattingDocument('inherited'),p=d.paragraphs[0]!,span=p.find('Alpha')[0]!,version=d.currentVersion();const first=p.effectiveRunFormatting();first[0]!.bold.value=false;first[0]!.paragraphStyleChain.push('Fake');first[0]!.italic.contributions[1]!.value=false;
 const fresh=p.effectiveRunFormatting();expect(fresh[0]!.bold.value).toBe(true);expect(fresh[0]!.paragraphStyleChain).toEqual(['Base','Child']);expect(d.currentVersion()).toBe(version);span.replace('Changed');const q=await Document.open(d.package.toBytes());expect(q.paragraphs[0]!.effectiveRunFormatting()[0]!.text).toBe('Changed');
});
test('style definitions are read live while raw main-document edits make handles stale',async()=>{
 const d=await formattingDocument('inherited'),p=d.paragraphs[0]!,part='word/styles.xml';
 const style=new TextDecoder().decode(d.package.get(part));d.package.setPart(part,encoder.encode(style.replace('<w:b/>','<w:b w:val="0"/>')));const before=d.package.toBytes();expect(p.effectiveRunFormatting()[0]!.bold.value).toBe(false);expect(d.package.toBytes()).toEqual(before);
 const main='word/document.xml';d.package.setPart(main,encoder.encode(new TextDecoder().decode(d.package.get(main))+' '));const raw=d.package.toBytes();expect(()=>p.effectiveRunFormatting()).toThrow('stale');expect(d.package.toBytes()).toEqual(raw);
});
test('UTF16 LE and BE preserve encoding/BOM and inspect reopened style definitions',async()=>{
 for(const le of [true,false]){const d=await formattingDocument('inherited');for(const name of ['word/document.xml','word/styles.xml']){let xml=new TextDecoder().decode(d.package.get(name));xml=xml.startsWith('<?xml')?xml.replace('encoding="UTF-8"','encoding="UTF-16"'):'<?xml version="1.0" encoding="UTF-16"?>'+xml;const bytes=new Uint8Array(xml.length*2+2),view=new DataView(bytes.buffer);view.setUint16(0,0xfeff,le);for(let i=0;i<xml.length;i++)view.setUint16(2+2*i,xml.charCodeAt(i),le);d.package.setPart(name,bytes);}const q=await Document.open(d.package.toBytes()),before=q.package.toBytes();expect(q.paragraphs[0]!.effectiveRunFormatting()[0]!.bold.value).toBe(true);expect(q.package.toBytes()).toEqual(before);for(const name of ['word/document.xml','word/styles.xml'])expect([...q.package.get(name)!.slice(0,2)]).toEqual(le?[255,254]:[254,255]);}
});
test('selected ancestry refuses missing bases, links, wrong namespaces and duplicate property containers',async()=>{
 for(const update of [(s:string)=>s.replace('w:val="Base"','w:val="Missing"'),(s:string)=>s.replace('<w:basedOn','<w:link w:val="Linked"/><w:basedOn'),(s:string)=>s.replace('<w:b/>','<x:b xmlns:x="urn:wrong"/>'),(s:string)=>s.replace('<w:b/>','<w:b val="0"/>'),(s:string)=>s.replace('<w:rPr><w:b/></w:rPr>','<w:rPr/><w:rPr/>'),(s:string)=>s.replace('<w:b/>','<w:b w:val="TRUE"/>')]){
 const d=await changed('inherited','word/styles.xml',update),before=d.package.toBytes();expect(()=>d.paragraphs[0]!.effectiveRunFormatting()).toThrow();expect(d.package.toBytes()).toEqual(before);}
});
test('unselected style ancestry is not evaluated but registry IDs remain unambiguous',async()=>{
 const d=await changed('inherited','word/styles.xml',s=>s.replace('</w:styles>','<w:style w:type="paragraph" w:styleId="Unused"><w:basedOn w:val="Unused"/><w:rPr><w:b w:val="invalid"/></w:rPr></w:style></w:styles>'));
 expect(d.paragraphs[0]!.effectiveRunFormatting()[0]!.paragraphStyleChain).toEqual(['Base','Child']);
});
test('protected documents remain readable without changing settings or paragraph handles',async()=>{
 const d=await formattingDocument('inherited'),pkg=await OpcPackage.open(d.package.toBytes());const {addPart,addRelationship}=await import('../../src/opc/index.ts');addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,'word/document.xml','http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');const q=await Document.open(pkg.toBytes()),before=q.package.toBytes();expect(q.paragraphs[0]!.effectiveRunFormatting()[0]!.bold.value).toBe(true);expect(q.package.toBytes()).toEqual(before);
});
test('disk save and reopen produce identical effective run formatting',async()=>{
 const {mkdtemp,rm}=await import('node:fs/promises'),{tmpdir}=await import('node:os'),{join}=await import('node:path'),root=await mkdtemp(join(tmpdir(),'bun-effective-'));
 try{const d=await formattingDocument('multiple-runs'),expected=d.paragraphs[0]!.effectiveRunFormatting(),path=join(root,'effective.docx');await d.save(path);const q=await Document.open(path);expect(q.paragraphs[0]!.effectiveRunFormatting()).toEqual(expected);}finally{await rm(root,{recursive:true,force:true});}
});
test('complex-script text and unsupported property contexts refuse rather than guess flags',async()=>{
 for(const update of [(s:string)=>s.replace('Alpha','שלום'),(s:string)=>s.replace('<w:r>','<w:r><w:rPr><w:rtl/></w:rPr>'),(s:string)=>s.replace('<w:pPr>','<w:pPr><w:bidi/>'),(s:string)=>s.replace('<w:r>','<w:r><w:rPr><w:bCs/></w:rPr>')]){const d=await changed('inherited','word/document.xml',update),before=d.package.toBytes();expect(()=>d.paragraphs[0]!.effectiveRunFormatting()).toThrow();expect(d.package.toBytes()).toEqual(before);}
});
test('raw root relationship changes cannot repoint the document behind a formatting handle',async()=>{
 const d=await formattingDocument('inherited'),p=d.paragraphs[0]!;const root='_rels/.rels',xml=new TextDecoder().decode(d.package.get(root));d.package.setPart(root,encoder.encode(xml.replace('word/document.xml','word/styles.xml')));const before=d.package.toBytes();expect(()=>p.effectiveRunFormatting()).toThrow();expect(d.package.toBytes()).toEqual(before);
});
