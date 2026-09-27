import { expect, test } from 'bun:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { Document, type RunFormattingPatch } from '../../src/index.ts';
import { OpcPackage, addPart, addRelationship } from '../../src/opc/index.ts';
import { parseXml, elements, attribute } from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(pr=''){
 const d=Document.create();d.addParagraph('Alpha');const pkg=await OpcPackage.open(d.package.toBytes()),xml=pkg.text(pkg.mainPart());expect(xml.includes('<w:r>')).toBe(true);pkg.set(pkg.mainPart(),xml.replace('<w:r>','<w:r>'+pr));return Document.open(pkg.toBytes());
}

test('fontName writes both Latin slots and reads back through saved archive without altering text or other parts',async()=>{
 const d=await fixture(),before=d.package.parts;expect(d.paragraphs[0]!.directFontNames()).toEqual([null]);
 expect(d.paragraphs[0]!.setRunFormatting({fontName:'Times New Roman'})).toEqual({changedRuns:1});
 const p=await Document.open(d.package.toBytes());expect(p.paragraphs[0]!.directFontNames()).toEqual(['Times New Roman']);expect(p.paragraphs[0]!.text).toBe('Alpha');
 const font=elements(parseXml(new TextDecoder().decode(p.package.get('word/document.xml'))),'rFonts',W)[0]!;
 expect(attribute(font,'ascii',W)).toBe('Times New Roman');expect(attribute(font,'hAnsi',W)).toBe('Times New Roman');expect(attribute(font,'eastAsia',W)).toBeUndefined();expect(attribute(font,'cs',W)).toBeUndefined();
 for(const[n,b]of before)if(n!=='word/document.xml')expect(p.package.get(n)).toEqual(b);
});

test('font override is ordered before flags and retains unrelated run property bytes',async()=>{
 const d=await fixture('<w:rPr><w:b/><w:color w:val = \'112233\'/><w:lang w:val="en-US"/></w:rPr>');
 d.paragraphs[0]!.setRunFormatting({highlight:'yellow',fontSizePt:14,fontName:'Arial'});
 const xml=new TextDecoder().decode(d.package.get('word/document.xml'));expect(xml).toContain('<w:b/>');expect(xml).toContain('<w:color w:val = \'112233\'/>');
 expect(elements(parseXml(xml),'rPr',W)[0]!.children.map(n=>n.localName)).toEqual(['rFonts','b','color','sz','highlight','lang']);
});

test('same decoded name is a lexical no-op and null removes the complete bounded font leaf',async()=>{
 const leaf='<w:rFonts w:ascii = \'A&amp;B\' w:hAnsi="A&amp;B"/>',d=await fixture('<w:rPr>'+leaf+'<w:i/></w:rPr>'),p=d.paragraphs[0]!,before=d.package.toBytes();
 expect(p.setRunFormatting({fontName:'A&B'})).toEqual({changedRuns:0});expect(p.text).toBe('Alpha');expect(d.package.toBytes()).toEqual(before);
 const values=p.directFontNames();values[0]='changed';expect(p.directFontNames()).toEqual(['A&B']);
 expect(p.setRunFormatting({fontName:null})).toEqual({changedRuns:1});expect(d.paragraphs[0]!.directFontNames()).toEqual([null]);expect(()=>p.directFontNames()).toThrow(expect.objectContaining({code:'docx-stale-paragraph'}));
 const after=d.package.toBytes();expect(d.paragraphs[0]!.setRunFormatting({fontName:null})).toEqual({changedRuns:0});expect(d.package.toBytes()).toEqual(after);
});

test('Unicode and XML metacharacters are escaped without changing a bounded font name',async()=>{
 const value='雪 & "quotes" <serif> \'face\'',d=await fixture();d.paragraphs[0]!.setRunFormatting({fontName:value});
 const xml=new TextDecoder().decode(d.package.get('word/document.xml'));expect(xml).toContain('&amp;');expect(xml).toContain('&quot;');expect(xml).toContain('&lt;');
 expect((await Document.open(d.package.toBytes())).paragraphs[0]!.directFontNames()).toEqual([value]);
 const edge=await fixture();edge.paragraphs[0]!.setRunFormatting({fontName:'a'.repeat(255)});expect(edge.paragraphs[0]!.directFontNames()[0]!.length).toBe(255);
});

test('invalid names or executable patches refuse without changing package bytes',async()=>{
 const d=await fixture(),before=d.package.toBytes();let called=false;
 const values:unknown[]=['',' ',' leading','trailing ','a\tb','a\nb','a\u0000b','a\u007fb','a\u0085b','a\u009fb','\ud800','a'.repeat(256),12,false];
 for(const value of values){expect(()=>d.paragraphs[0]!.setRunFormatting({fontName:value} as RunFormattingPatch)).toThrow(expect.objectContaining({code:'docx-format-argument'}));expect(d.package.toBytes()).toEqual(before);}
 const patch=Object.defineProperty({},'fontName',{enumerable:true,get(){called=true;return 'Arial';}});expect(()=>d.paragraphs[0]!.setRunFormatting(patch)).toThrow();expect(called).toBe(false);
});

test('theme, script, hint, incomplete and unequal font slots refuse even for removal or identical request',async()=>{
 const leaves=[
  '<w:rFonts w:ascii="Arial"/>','<w:rFonts w:hAnsi="Arial"/>','<w:rFonts/>',
  '<w:rFonts w:ascii="Arial" w:hAnsi="Calibri"/>',
  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:asciiTheme="majorAscii"/>',
  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hAnsiTheme="majorHAnsi"/>',
  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Font"/>',
  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Font"/>',
  '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/>',
  '<w:rFonts ascii="Arial" w:hAnsi="Arial"/>',
 ];
 for(const leaf of leaves){const d=await fixture('<w:rPr>'+leaf+'</w:rPr>'),before=d.package.toBytes();
  expect(()=>d.paragraphs[0]!.directFontNames()).toThrow(expect.objectContaining({code:'docx-format-unsupported'}));
  for(const value of ['Arial',null]){expect(()=>d.paragraphs[0]!.setRunFormatting({fontName:value})).toThrow();expect(d.package.toBytes()).toEqual(before);}
  d.paragraphs[0]!.setRunFormatting({bold:true});expect(new TextDecoder().decode(d.package.get('word/document.xml'))).toContain(leaf);
 }
});

test('late multi-run font metadata refusal leaves earlier runs and handles untouched',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('</w:r>','</w:r><w:r><w:rPr><w:rFonts w:asciiTheme="majorAscii"/></w:rPr><w:t>Beta</w:t></w:r>'));
 const doc=await Document.open(pkg.toBytes()),before=doc.package.toBytes(),held=doc.paragraphs[0]!;expect(()=>held.setRunFormatting({fontName:'Arial',bold:true})).toThrow();expect(doc.package.toBytes()).toEqual(before);expect(held.text).toBe('AlphaBeta');
});

test('namespace aliases and UTF-16LE disk save retain matching Latin names and unrelated payloads',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()).replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('UTF-8','UTF-16');
 const bytes=new Uint8Array(2+text.length*2);bytes[0]=255;bytes[1]=254;const view=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);
 const doc=await Document.open(pkg.toBytes());doc.paragraphs[0]!.setRunFormatting({fontName:'Courier New'});const root=await mkdtemp(join(tmpdir(),'run-font-'));
 try{const path=join(root,'font.docx');await doc.save(path);const reopened=await Document.open(path);expect(reopened.paragraphs[0]!.directFontNames()).toEqual(['Courier New']);expect([...reopened.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(reopened.package.get(n)).toEqual(pkg.get(n));}finally{await rm(root,{recursive:true,force:true});}
});

test('protected writes, stale source and serialization failure preserve bytes and model state',async()=>{
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');
 const locked=await Document.open(pkg.toBytes()),lb=locked.package.toBytes();expect(()=>locked.paragraphs[0]!.setRunFormatting({fontName:'Arial'})).toThrow();expect(locked.package.toBytes()).toEqual(lb);
 const held=d.paragraphs[0]!,before=d.package.toBytes(),opc=(d as unknown as {opcPackage:OpcPackage}).opcPackage,original=opc.toBytes.bind(opc);opc.toBytes=()=>{throw Error('injected serialization');};try{expect(()=>held.setRunFormatting({fontName:'Arial'})).toThrow('injected serialization');}finally{opc.toBytes=original;}
 expect(d.package.toBytes()).toEqual(before);expect(held.directFontNames()).toEqual([null]);
 d.package.setPart('word/document.xml',new TextEncoder().encode(new TextDecoder().decode(d.package.get('word/document.xml')).replace('Alpha','Beta')));const changed=d.package.toBytes();expect(()=>held.directFontNames()).toThrow();expect(()=>held.setRunFormatting({fontName:'Arial'})).toThrow();expect(d.package.toBytes()).toEqual(changed);
});

test('six shared font-name getter cases execute and corrupted slot data fails real predicates',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts');const {bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/run-font-name.ts');
 const path='workflows/docx/run-formatting.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inventory=(text:string)=>({root:'.',features:[selectSharedScenarios(path,text,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(6),steps:count(18)}});
 const good=await executeAcceptance(inventory(source),bindings,'run-font-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(6);expect(good.counts.cases.planned).toBeGreaterThan(0);
 const wrong=await executeAcceptance(inventory(source.replace('its font-name getter equals <font>','its font-name getter equals Wrong')),bindings,'run-font-expected-control');expect(wrong.counts.cases.failed).toBe(6);expect(wrong.counts.steps.failed).toBe(6);expect(wrong.counts.steps.undefined).toBe(0);expect(wrong.counts.steps.ambiguous).toBe(0);
 const corrupted=bindings.map(b=>b.pattern.test('its font name is set to Arial')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);(c.state as {document:Document}).document.paragraphs[0]!.setRunFormatting({fontName:'Wrong'});}}:b);
 const bad=await executeAcceptance(inventory(source),corrupted,'run-font-data-control');expect(bad.counts.cases.failed).toBe(6);expect(bad.counts.steps.failed).toBe(6);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);
});

test('font changes in table paragraphs preserve grid and staged-write rollback retains fresh handles',()=>{
 const d=Document.create();d.addTable(1,2);d.tables[0]!.cell(0,0).text='Cell';d.tables[0]!.cell(0,1).text='Other';const table=d.tables[0]!,p=d.paragraphs[0]!;
 p.setRunFormatting({fontName:'Arial'});expect(table.rows).toBe(1);expect(table.columns).toBe(2);expect(table.cell(0,0).text).toBe('Cell');expect(table.cell(0,1).text).toBe('Other');expect(()=>p.directFontNames()).toThrow();
 const held=d.paragraphs[0]!,before=d.package.toBytes(),opc=(d as unknown as {opcPackage:OpcPackage}).opcPackage,original=opc.set.bind(opc);opc.set=(name,value)=>{original(name,value);throw Error('injected write');};
 try{expect(()=>held.setRunFormatting({fontName:'Verdana'})).toThrow('injected write');}finally{opc.set=original;}
 expect(d.package.toBytes()).toEqual(before);expect(held.directFontNames()).toEqual(['Arial']);
});
