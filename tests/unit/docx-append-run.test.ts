import {sharedScenarios} from '../helpers/shared-scenarios.ts';
import {expect,test} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {tmpdir} from 'node:os';import {join} from 'node:path';
import {Document,type RunFormattingPatch} from '../../src/index.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {parseXml,elements,applyEdits} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(content:string){const d=Document.create();d.addParagraph('Alpha');const pkg=await OpcPackage.open(d.package.toBytes()),xml=pkg.text(pkg.mainPart()),p=elements(parseXml(xml),'p',W)[0]!;pkg.set(pkg.mainPart(),applyEdits(xml,[{start:p.openEnd,end:p.closeStart,value:content}]));return Document.open(pkg.toBytes());}

test('three appended runs remain distinct and concatenate through returned fresh paragraph handles',async()=>{
 const d=Document.create(),first=d.addParagraph('');const second=first.appendRun('Hello '),third=second.appendRun('World'),last=third.appendRun('!');
 expect(last.text).toBe('Hello World!');expect(d.paragraphs).toHaveLength(1);expect(last).toBe(d.paragraphs[0]!);
 expect(()=>first.text).toThrow(expect.objectContaining({code:'docx-stale-paragraph'}));expect(()=>second.appendRun('stale')).toThrow();
 expect(last.directRunFlags()).toHaveLength(3);
 const reopened=await Document.open(d.package.toBytes());expect(reopened.paragraphs[0]!.text).toBe('Hello World!');expect(reopened.paragraphs[0]!.directRunFlags()).toHaveLength(3);
});

test('per-run authoring preserves selected formatting after real path save and reopen',async()=>{
 const d=Document.create();let p=d.addParagraph('');p=p.appendRun('Bold ',{bold:true});p=p.appendRun('Italic ',{italic:true});p=p.appendRun('Colored',{color:'FF0000',fontSizePt:14,fontName:'Arial'});
 expect(p.text).toBe('Bold Italic Colored');const root=await mkdtemp(join(tmpdir(),'docx-append-run-'));
 try{const path=join(root,'runs.docx');await d.save(path);const after=await Document.open(path),q=after.paragraphs[0]!;expect(q.directRunFlags().map(v=>[v.bold,v.italic])).toEqual([[true,null],[null,true],[null,null]]);expect(q.directRunAppearance().map(v=>v.color)).toEqual([null,null,'FF0000']);expect(q.directFontSizes()).toEqual([null,null,14]);expect(q.directFontNames()).toEqual([null,null,'Arial']);expect(q.text).toBe('Bold Italic Colored');}finally{await rm(root,{recursive:true,force:true});}
});

test('append retains existing paragraph/run fragments, sibling paragraphs and unrelated member bytes',async()=>{
 const content='<w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/><w:color w:val = \'123456\'/></w:rPr><w:t>Alpha</w:t></w:r>';
 const d=await fixture(content);d.addParagraph('Untouched');const before=d.package.parts,xmlBefore=new TextDecoder().decode(before.get('word/document.xml'));const sibling=elements(parseXml(xmlBefore),'p',W)[1]!,siblingXml=xmlBefore.slice(sibling.start,sibling.end);
 d.paragraphs[0]!.appendRun(' & <雪> ',{italic:true});const xml=new TextDecoder().decode(d.package.get('word/document.xml'));expect(xml).toContain(content);expect(xml).toContain(siblingXml);expect(xml).toContain('xml:space="preserve"');expect(xml).toContain('&amp; &lt;雪&gt;');expect(d.paragraphs[0]!.text).toBe('Alpha & <雪> ');
 for(const[n,b]of before)if(n!=='word/document.xml')expect(d.package.get(n)).toEqual(b);
});

test('self-closing paragraphs expand with namespace safety and empty text still appends one run',async()=>{
 const d=Document.create(),pkg=await OpcPackage.open(d.package.toBytes());pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('<w:body>',`<w:body><q:p xmlns:q="${W}" xmlns:w="urn:foreign" />`));const doc=await Document.open(pkg.toBytes());
 let p=doc.paragraphs[0]!.appendRun('one',{bold:true});expect(p.text).toBe('one');p=p.appendRun('');expect(p.text).toBe('one');expect(p.directRunFlags()).toHaveLength(2);
 const xml=parseXml(new TextDecoder().decode(doc.package.get('word/document.xml')));expect(elements(xml,'r',W)).toHaveLength(2);expect(elements(xml,'r','urn:foreign')).toHaveLength(0);expect(p.directRunFlags()[0]!.bold).toBe(true);
});

test('invalid text, executable patches and conflicting formatting refuse before mutation',()=>{
 const d=Document.create(),p=d.addParagraph('Alpha'),before=d.package.toBytes();let called=false;
 for(const text of [null,1,'bad\u0000','\ud800','a\tb','a\nb','a\rb','x'.repeat(1024*1024+1)]){expect(()=>p.appendRun(text as string)).toThrow();expect(d.package.toBytes()).toEqual(before);}
 for(const patch of [{bold:'yes'},{strike:true,doubleStrike:true},{fontName:' '},{color:'red'}]){expect(()=>p.appendRun('good',patch as RunFormattingPatch)).toThrow();expect(d.package.toBytes()).toEqual(before);}
 const patch=Object.defineProperty({},'bold',{enumerable:true,get(){called=true;return true;}});expect(()=>p.appendRun('x',patch)).toThrow();expect(called).toBe(false);expect(p.text).toBe('Alpha');
});

test('unsupported, revised and lexical target topology refuses without silently flattening',async()=>{
 for(const content of ['<w:hyperlink><w:r><w:t>Alpha</w:t></w:r></w:hyperlink>','<w:r><w:fldChar w:fldCharType="begin"/></w:r>','<w:r><w:t>Alpha</w:t></w:r><!--keep-->','<w:pPr><w:pPrChange/></w:pPr><w:r><w:t>Alpha</w:t></w:r>','<w:pPr><w:jc w:val="left"/><w:keepLines/></w:pPr><w:r><w:t>Alpha</w:t></w:r>']){
  const d=await fixture(content),before=d.package.toBytes();expect(()=>d.paragraphs[0]!.appendRun('No')).toThrow();expect(d.package.toBytes()).toEqual(before);
 }
});

test('document protection and external XML changes refuse held append handles',async()=>{
 const d=Document.create();d.addParagraph('Alpha');const pkg=await OpcPackage.open(d.package.toBytes());addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');
 const locked=await Document.open(pkg.toBytes()),lb=locked.package.toBytes();expect(()=>locked.paragraphs[0]!.appendRun('x')).toThrow();expect(locked.package.toBytes()).toEqual(lb);
 const held=d.paragraphs[0]!;d.package.setPart('word/document.xml',new TextEncoder().encode(new TextDecoder().decode(d.package.get('word/document.xml')).replace('Alpha','Beta')));const before=d.package.toBytes();expect(()=>held.appendRun('x')).toThrow(expect.objectContaining({code:'docx-stale-paragraph'}));expect(d.package.toBytes()).toEqual(before);
});

test('staged write and serialization failure preserve paragraph/cell handles and package',()=>{
 for(const stage of ['set','toBytes'] as const){const d=Document.create();d.addTable(1,1);d.tables[0]!.cell(0,0).text='Alpha';const p=d.paragraphs[0]!,cell=d.tables[0]!.cell(0,0),before=d.package.toBytes(),pkg=(d as unknown as{opcPackage:OpcPackage}).opcPackage,old=pkg[stage].bind(pkg);
  if(stage==='set')pkg.set=(n,v)=>{(old as OpcPackage['set'])(n,v);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialization');};
  try{expect(()=>p.appendRun('Beta',{bold:true})).toThrow('injected');}finally{if(stage==='set')pkg.set=old as OpcPackage['set'];else pkg.toBytes=old as OpcPackage['toBytes'];}
  expect(d.package.toBytes()).toEqual(before);expect(p.text).toBe('Alpha');expect(cell.text).toBe('Alpha');expect(p.appendRun('Beta').text).toBe('AlphaBeta');
 }
});

test('table append preserves grid and expires cell snapshots, with UTF-16 payload retained',async()=>{
 const d=Document.create();d.addTable(1,2);d.tables[0]!.cell(0,0).text='Alpha';d.tables[0]!.cell(0,1).text='Other';const pkg=await OpcPackage.open(d.package.toBytes()),text=pkg.text(pkg.mainPart()).replace('UTF-8','UTF-16'),bytes=new Uint8Array(2+text.length*2);bytes[0]=255;bytes[1]=254;const v=new DataView(bytes.buffer);for(let i=0;i<text.length;i++)v.setUint16(2+2*i,text.charCodeAt(i),true);pkg.set(pkg.mainPart(),bytes);
 const doc=await Document.open(pkg.toBytes()),table=doc.tables[0]!,cell=table.cell(0,0);doc.paragraphs[0]!.appendRun(' + 雪',{underline:'single'});expect(table.rows).toBe(1);expect(table.columns).toBe(2);expect(()=>cell.text).toThrow();expect(table.cell(0,0).text).toBe('Alpha + 雪');expect(table.cell(0,1).text).toBe('Other');
 const after=await Document.open(doc.package.toBytes());expect([...after.package.get('word/document.xml')!.slice(0,2)]).toEqual([255,254]);expect(after.paragraphs[0]!.text).toBe('Alpha + 雪');expect(after.paragraphs[0]!.directRunAppearance()[1]!.underline).toBe('single');
});

test('two shared run-authoring cases execute through real save/reopen and reject corrupted saved formatting',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts');const {bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/append-run.ts');
 const count=(n:number)=>({implemented:n,planned:0,total:n});
 const inv=async(change:(s:string)=>string=s=>s)=>({root:'.',features:await sharedScenarios(scenarioIds,change),counts:{features:count(2),scenarios:count(2),cases:count(2),steps:count(9)}});
 const good=await executeAcceptance(await inv(),bindings,'append-run-unit');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(2);expect(good.counts.cases.planned).toBe(0);
 const expected=await executeAcceptance(await inv(s=>s.replace('the third run reports colour FF0000, font size 14 and font Arial','the third run reports colour 000000, font size 12 and font Wrong')),bindings,'append-run-expectation');expect(expected.counts.cases.failed).toBe(1);expect(expected.counts.steps.failed).toBe(1);expect(expected.counts.steps.undefined).toBe(0);
 for(const patch of [{bold:false},{italic:false},{color:'000000'},{fontSizePt:12},{fontName:'Wrong'}]){
  const corrupted=bindings.map(b=>b.pattern.test('the document is saved and reopened')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);(c.state as {reopened:Document}).reopened.paragraphs[0]!.setRunFormatting(patch);}}:b);
  const bad=await executeAcceptance(await inv(),corrupted,'append-run-format-control');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);
 }
});

test('append defaults do not borrow direct formatting from previous runs and preserve explicit null',()=>{
 const d=Document.create();let p=d.addParagraph('Old',{bold:true});p=p.appendRun('Plain');p=p.appendRun('New',{bold:null,italic:true});expect(p.directRunFlags().map(f=>[f.bold,f.italic])).toEqual([[true,null],[null,null],[null,true]]);
 const before=d.package.toBytes(),held=p;expect(()=>p.appendRun('bad',{fontName:'Arial',strike:true,doubleStrike:true})).toThrow();expect(d.package.toBytes()).toEqual(before);expect(held.text).toBe('OldPlainNew');
});

test('run append enforces complete XML output bounds before mutating a nearly full package',async()=>{
 const d=Document.create();d.addParagraph('');const pkg=await OpcPackage.open(d.package.toBytes()),main=pkg.mainPart(),xml=pkg.text(main),padding='x'.repeat(8*1024*1024-xml.length-9);
 pkg.set(main,xml.replace('</w:body>','<!--'+padding+'--></w:body>'));const loaded=await Document.open(pkg.toBytes()),before=loaded.package.toBytes(),held=loaded.paragraphs[0]!;
 expect(()=>held.appendRun('a')).toThrow(expect.objectContaining({code:'XML_EDIT_UNSAFE'}));expect(loaded.package.toBytes()).toEqual(before);expect(held.text).toBe('');
});
