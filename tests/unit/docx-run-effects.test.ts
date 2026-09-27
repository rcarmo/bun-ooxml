import { expect, test } from 'bun:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { Document, type RunFormattingPatch } from '../../src/index.ts';
import { OpcPackage, addPart, addRelationship } from '../../src/opc/index.ts';
import { parseXml, elements, attribute } from '../../src/xml/index.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const effects = ['strike','doubleStrike','caps','smallCaps','outline','shadow','emboss','imprint','vanish'] as const;
async function fixture(rPr = '') {
 const d = Document.create(); d.addParagraph('Alpha'); const p = await OpcPackage.open(d.package.toBytes());
 const xml = p.text(p.mainPart()); expect(xml.includes('<w:r>')).toBe(true); p.set(p.mainPart(), xml.replace('<w:r>', '<w:r>' + rPr));
 return Document.open(p.toBytes());
}

test('direct boolean getter distinguishes absent/off/on and each new effect survives reopen', async () => {
 for (const effect of effects) {
  const d = await fixture(), original = d.package.parts;
  expect(d.paragraphs[0]!.directRunFlags()[0]![effect]).toBeNull();
  expect(d.paragraphs[0]!.setRunFormatting({ [effect]: true })).toEqual({ changedRuns: 1 });
  const reopened = await Document.open(d.package.toBytes()); expect(reopened.paragraphs[0]!.directRunFlags()[0]![effect]).toBe(true); expect(reopened.paragraphs[0]!.text).toBe('Alpha');
  for (const [name, bytes] of original) if(name !== 'word/document.xml') expect(reopened.package.get(name)).toEqual(bytes);
  reopened.paragraphs[0]!.setRunFormatting({ [effect]: false }); expect(reopened.paragraphs[0]!.directRunFlags()[0]![effect]).toBe(false);
  reopened.paragraphs[0]!.setRunFormatting({ [effect]: null }); expect(reopened.paragraphs[0]!.directRunFlags()[0]![effect]).toBeNull();
 }
});

test('valid compound edits follow schema order regardless of input field order and retain unrelated properties', async () => {
 const d = await fixture('<w:rPr><w:rFonts w:ascii="Arial"/><w:color w:val="112233"/><w:u w:val="single"/></w:rPr>');
 d.paragraphs[0]!.setRunFormatting({ fontSizePt: 14, vanish: true, shadow: true, strike: true, caps: true, italic: false, bold: true });
 const xml = parseXml(new TextDecoder().decode(d.package.get('word/document.xml'))), pr = elements(xml, 'rPr', W)[0]!;
 expect(pr.children.map(n => n.localName)).toEqual(['rFonts','b','i','caps','strike','shadow','vanish','color','sz','u']);
 expect(attribute(pr.children[0]!, 'ascii', W)).toBe('Arial'); expect(attribute(elements(xml,'color',W)[0]!, 'val', W)).toBe('112233');
 expect(d.paragraphs[0]!.directRunFlags()[0]).toEqual({bold:true,italic:false,caps:true,smallCaps:null,strike:true,doubleStrike:null,outline:null,shadow:true,emboss:null,imprint:null,vanish:true});
});

test('same-state/no-op preserves lexical spelling and fresh handle; real effect edit expires old handles', async () => {
 const d = await fixture('<w:rPr><w:strike w:val="on"/><w:shadow w:val="off"/></w:rPr>'), p=d.paragraphs[0]!, before=d.package.toBytes();
 expect(p.setRunFormatting({ strike:true, shadow:false })).toEqual({changedRuns:0}); expect(d.package.toBytes()).toEqual(before); expect(p.text).toBe('Alpha');
 const returned=p.directRunFlags();returned[0]!.strike=false;expect(p.directRunFlags()[0]!.strike).toBe(true);
 p.setRunFormatting({ strike:null });expect(()=>p.directRunFlags()).toThrow(expect.objectContaining({code:'docx-stale-paragraph'}));
});

test('spec-incompatible direct effect pairs refuse atomically and null can remove a conflicting property', async () => {
 const pairs = [['strike','doubleStrike'],['caps','smallCaps'],['emboss','imprint'],['emboss','outline'],['imprint','outline'],['emboss','shadow'],['imprint','shadow']] as const;
 for(const [a,b] of pairs){const d=await fixture(),before=d.package.toBytes();
  expect(()=>d.paragraphs[0]!.setRunFormatting({[a]:true,[b]:true})).toThrow(expect.objectContaining({code:'docx-format-unsupported'}));expect(d.package.toBytes()).toEqual(before);
  d.paragraphs[0]!.setRunFormatting({[a]:true});const current=d.package.toBytes();
  expect(()=>d.paragraphs[0]!.setRunFormatting({[b]:true})).toThrow();expect(d.package.toBytes()).toEqual(current);
  d.paragraphs[0]!.setRunFormatting({[a]:null,[b]:true});expect(d.paragraphs[0]!.directRunFlags()[0]![a]).toBeNull();expect(d.paragraphs[0]!.directRunFlags()[0]![b]).toBe(true);
 }
});

test('malformed effect metadata, revisions and unsupported run topology refuse even for no-op', async () => {
 for(const pr of ['<w:rPr><w:strike val="true"/></w:rPr>','<w:rPr><w:shadow w:val="maybe"/></w:rPr>','<w:rPr><w:caps/><w:caps/></w:rPr>','<w:rPr><w:vanish/><w:strike/></w:rPr>','<w:rPr><w:rPrChange/></w:rPr>','<w:rPr><!--keep--><w:strike/></w:rPr>','<w:rPr><w:strike/><w:dstrike/></w:rPr>']) {
  const d=await fixture(pr),before=d.package.toBytes();expect(()=>d.paragraphs[0]!.setRunFormatting({strike:true})).toThrow();expect(d.package.toBytes()).toEqual(before);
 }
});

test('patch getters, prototypes, symbols and invalid values refuse before running user accessors', async () => {
 const d=await fixture(),before=d.package.toBytes();let read=false;
 const patches:unknown[]=[{strike:1},{shadow:'true'},{bold:true,vanish:'bad'},Object.create({bold:true}),{[Symbol('x')]:true}];
 patches.push(Object.defineProperty({},'strike',{enumerable:true,get(){read=true;return true;}}));
 for(const patch of patches){expect(()=>d.paragraphs[0]!.setRunFormatting(patch as RunFormattingPatch)).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(read).toBe(false);
});

test('namespace aliases and UTF-16 path save preserve effects and nonedited member payloads', async () => {
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());
 const text=pkg.text(pkg.mainPart()).replaceAll('xmlns:w=','xmlns:q=').replaceAll('w:','q:').replace('UTF-8','UTF-16');
 const b=new Uint8Array(2+text.length*2);b[0]=255;b[1]=254;const view=new DataView(b.buffer);for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),true);pkg.set(pkg.mainPart(),b);
 const doc=await Document.open(pkg.toBytes());doc.paragraphs[0]!.setRunFormatting({shadow:true,strike:true});const root=await mkdtemp(join(tmpdir(),'docx-effects-'));
 try{const path=join(root,'effects.docx');await doc.save(path);const reopened=await Document.open(path);expect(reopened.paragraphs[0]!.directRunFlags()[0]!.shadow).toBe(true);expect(reopened.paragraphs[0]!.directRunFlags()[0]!.strike).toBe(true);expect([...reopened.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);for(const n of pkg.names())if(n!==pkg.mainPart())expect(reopened.package.get(n)).toEqual(pkg.get(n));}finally{await rm(root,{recursive:true,force:true});}
});

test('protected requests and staged serialization failures preserve bytes and handles', async () => {
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');
 const locked=await Document.open(pkg.toBytes()),beforeLocked=locked.package.toBytes();expect(()=>locked.paragraphs[0]!.setRunFormatting({strike:true})).toThrow();expect(locked.package.toBytes()).toEqual(beforeLocked);
 const opc=(d as unknown as {opcPackage:OpcPackage}).opcPackage,save=opc.toBytes.bind(opc),before=d.package.toBytes(),held=d.paragraphs[0]!;
 opc.toBytes=()=>{throw Error('injected serialization');};try{expect(()=>held.setRunFormatting({strike:true})).toThrow('injected serialization');}finally{opc.toBytes=save;}
 expect(d.package.toBytes()).toEqual(before);expect(held.directRunFlags()[0]!.strike).toBeNull();expect(held.text).toBe('Alpha');
});

test('late multi-run conflict refuses the whole paragraph and requested removals can switch effects', async () => {
 const d=await fixture(),pkg=await OpcPackage.open(d.package.toBytes());
 pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('</w:r>','</w:r><w:r><w:rPr><w:dstrike/></w:rPr><w:t>Beta</w:t></w:r>'));
 const doc=await Document.open(pkg.toBytes()),before=doc.package.toBytes();
 expect(()=>doc.paragraphs[0]!.setRunFormatting({strike:true})).toThrow();expect(doc.package.toBytes()).toEqual(before);expect(doc.paragraphs[0]!.text).toBe('AlphaBeta');
 expect(doc.paragraphs[0]!.setRunFormatting({doubleStrike:null,strike:true})).toEqual({changedRuns:2});
 expect(doc.paragraphs[0]!.directRunFlags().map(f=>[f.strike,f.doubleStrike])).toEqual([[true,null],[true,null]]);
});

test('explicit false also participates in conservative conflict refusal and getter rejects invalid pairs', async () => {
 const d=await fixture('<w:rPr><w:caps w:val="0"/></w:rPr>'),before=d.package.toBytes();
 expect(()=>d.paragraphs[0]!.setRunFormatting({smallCaps:false})).toThrow();expect(d.package.toBytes()).toEqual(before);
 const invalid=await fixture('<w:rPr><w:caps w:val="0"/><w:smallCaps w:val="0"/></w:rPr>');
 expect(()=>invalid.paragraphs[0]!.directRunFlags()).toThrow();
});

test('effect getters reject stale external XML and formatting preserves live table grid', async () => {
 const d=Document.create();d.addTable(1,1);d.tables[0]!.cell(0,0).text='Cell';const table=d.tables[0]!,held=d.paragraphs[0]!;
 held.setRunFormatting({vanish:true});expect(table.rows).toBe(1);expect(table.columns).toBe(1);expect(table.cell(0,0).text).toBe('Cell');expect(()=>held.directRunFlags()).toThrow();
 const p=d.paragraphs[0]!,xml=new TextDecoder().decode(d.package.get('word/document.xml'));d.package.setPart('word/document.xml',new TextEncoder().encode(xml.replace('Cell','Changed')));
 const before=d.package.toBytes();expect(()=>p.directRunFlags()).toThrow();expect(()=>p.setRunFormatting({shadow:true})).toThrow();expect(d.package.toBytes()).toEqual(before);
});

test('five shared direct flag cases execute but spec-incompatible all-effects scenario stays planned', async () => {
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts');
 const {bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/run-effects.ts');
 const path='workflows/docx/run-formatting.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inventory={root:'.',features:[selectSharedScenarios(path,source,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(5),steps:count(15)}};
 const good=await executeAcceptance(inventory,bindings,'run-effects-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(5);
 expect(good.features[0]!.scenarios.find(s=>s.scenarioId==='@id-docx-go-run-effects-getters')!.result).toBe('planned');
 const corrupt=bindings.map(b=>b.pattern.test('bold is set to true, italic to true and strike to true')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);(c.state as {document:Document}).document.paragraphs[0]!.setRunFormatting({strike:captures[2]!=='true'});}}:b);
 const bad=await executeAcceptance(inventory,corrupt,'run-effects-negative');expect(bad.counts.cases.failed).toBe(5);expect(bad.counts.steps.failed).toBe(5);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);
});
