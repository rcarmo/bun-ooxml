import {test,expect} from 'bun:test';
import {Presentation,OpcPackage} from '../../src/index.ts';
import {fixturePath,F} from '../../scripts/fixture-inputs.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
const notes='ppt/notesSlides/notesSlide1.xml',slide='ppt/slides/slide1.xml';
const source=async()=>Uint8Array.from(await Bun.file(fixturePath(F.goSlides.notes)).bytes());
async function deck(paragraph?:string){const p=await OpcPackage.open(await source());if(paragraph){const text=p.text(notes),old='<a:p><a:r><a:t>Remember to emphasize the Gothic elements</a:t></a:r></a:p>';expect(text).toContain(old);p.set(notes,text.replace(old,paragraph));}return Presentation.open(p.toBytes());}
test('existing notes replacement changes only its notes payload and reopens exact text',async()=>{
 const bytes=await source(),p=await Presentation.open(bytes),s=p.slides[0]!,target=s.inspectNotesText();expect(target.text).toBe('Remember to emphasize the Gothic elements');
 const before=p.package.text(notes);expect(s.replaceNotesAt(target,'Updated speaker notes')).toEqual({changedParts:[notes]});expect(p.package.text(notes)).toBe(before.replace('Remember to emphasize the Gothic elements','Updated speaker notes'));
 const saved=p.package.toBytes(),after=await OpcPackage.open(saved),original=await OpcPackage.open(bytes);expect(after.names()).toEqual(original.names());for(const name of original.names())if(name!==notes)expect(after.get(name)).toEqual(original.get(name));expect((await Presentation.open(saved)).slides[0]!.inspectNotesText().text).toBe('Updated speaker notes');
});
test('foreign, invalid and forged targets refuse atomically; exact no-op preserves bytes and target',async()=>{
 const bytes=await source(),p=await Presentation.open(bytes),other=await Presentation.open(bytes),s=p.slides[0]!,target=s.inspectNotesText();
 expect(()=>other.slides[0]!.replaceNotesAt(target,'foreign')).toThrow();expect(()=>s.replaceNotesAt({...target},'forged')).toThrow();
 for(const text of ['bad\0','two\tcolumns','bad\rline','\ud800']){expect(()=>s.replaceNotesAt(target,text)).toThrow();expect(p.package.toBytes()).toEqual(bytes);}
 expect(s.replaceNotesAt(target,target.text)).toEqual({changedParts:[]});expect(p.package.toBytes()).toEqual(bytes);s.replaceNotesAt(target,'fresh');expect(s.inspectNotesText().text).toBe('fresh');
});
test('changed notes stale other held targets and empty text can be refilled',async()=>{
 const p=await deck(),s=p.slides[0]!,first=s.inspectNotesText(),second=s.inspectNotesText();s.replaceNotesAt(first,'');expect(()=>s.replaceNotesAt(second,'stale')).toThrow('stale');expect(s.inspectNotesText().text).toBe('');s.replaceNotesAt(s.inspectNotesText(),'refilled');expect(s.inspectNotesText().text).toBe('refilled');
});
test('direct notes mutation and relationship mutation invalidate captured targets',async()=>{
 for(const path of [notes,'ppt/slides/_rels/slide1.xml.rels']){const p=await deck(),s=p.slides[0]!,t=s.inspectNotesText();p.package.set(path,p.package.text(path)+' ');const before=p.package.toBytes();expect(()=>s.replaceNotesAt(t,'stale')).toThrow('stale');expect(p.package.toBytes()).toEqual(before);}
});
test('self-closing leaf expands and edge whitespace writes XML space preserve',async()=>{
 const p=await deck('<a:p><a:r><a:rPr b="1"/><a:t/></a:r></a:p>'),s=p.slides[0]!;s.replaceNotesAt(s.inspectNotesText(),'filled');expect(s.inspectNotesText().text).toBe('filled');expect(p.package.text(notes)).toContain('<a:rPr b="1"/>');
 const spaced=await deck('<a:p><a:r><a:t xml:space="default">old</a:t></a:r></a:p>');spaced.slides[0]!.replaceNotesAt(spaced.slides[0]!.inspectNotesText(),' leading ');const leaf=elements(parseXml(spaced.package.text(notes)),'t','http://schemas.openxmlformats.org/drawingml/2006/main').find(n=>n.text===' leading ')!;expect(attribute(leaf,'space','http://www.w3.org/XML/1998/namespace')).toBe('preserve');
});
test('multiline uses only first-run template and preserves boundary empty paragraphs',async()=>{
 const p=await deck('<a:p><a:pPr algn="ctr"/><a:r><a:t>one</a:t></a:r><a:r><a:rPr b="1"/><a:t>two</a:t></a:r></a:p>'),s=p.slides[0]!,t=s.inspectNotesText(),before=p.package.get(notes);expect(s.replaceNotesAt(t,t.text)).toEqual({changedParts:[]});expect(p.package.get(notes)).toEqual(before);s.replaceNotesAt(t,'\n A&B \n雪\n');expect(s.inspectNotesText().text).toBe('\n A&B \n雪\n');expect(p.package.text(notes)).not.toContain('b="1"');
});
test('multiline copies selected paragraph and first-run property fragments',async()=>{
 const p=await deck('<a:p><a:pPr algn="ctr"><a:buChar char="•"/><a:defRPr sz="1200"/></a:pPr><a:r><a:rPr b="1"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="F&amp;F"/></a:rPr><a:t>old</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>'),s=p.slides[0]!;s.replaceNotesAt(s.inspectNotesText(),'a\nb');for(const token of ['val="112233"','typeface="F&amp;F"','lang="en-US"','char="•"'])expect(p.package.text(notes).split(token).length-1).toBe(2);
});
test('missing notes and unsupported field topology refuse without creating or rewriting parts',async()=>{
 const minimal=await Presentation.open(fixturePath(F.goSlides.minimal)),bytes=minimal.package.toBytes();expect(()=>minimal.slides[0]!.inspectNotesText()).toThrow();expect(minimal.package.toBytes()).toEqual(bytes);
 const p=await deck('<a:p><a:fld id="x" type="slidenum"><a:t>1</a:t></a:fld></a:p>'),before=p.package.toBytes();expect(()=>p.slides[0]!.inspectNotesText()).toThrow();expect(p.package.toBytes()).toEqual(before);
});

test('notes edits refuse text locks, extended properties and ambiguous graph ownership',async()=>{
 for(const variant of ['lock','extension','shared','fragment','protection','wrong-type','body-duplicate']){
  const p=await deck();if(variant==='lock')p.package.set(notes,p.package.text(notes).replaceAll('<a:spLocks noGrp="1"/>','<a:spLocks noTextEdit="1"/>'));
  if(variant==='extension')p.package.set(notes,p.package.text(notes).replace('<a:p><a:r>','<a:p><a:pPr><a:extLst/></a:pPr><a:r>'));
  if(variant==='shared'){const rp='ppt/slides/_rels/slide2.xml.rels';p.package.set(rp,p.package.text(rp).replace('notesSlide2.xml','notesSlide1.xml'));}
  if(variant==='fragment'){const rp='ppt/slides/_rels/slide1.xml.rels';p.package.set(rp,p.package.text(rp).replace('notesSlide1.xml','notesSlide1.xml#fragment'));}
  if(variant==='protection'){const main=p.package.mainPart();p.package.set(main,p.package.text(main).replace('</p:presentation>','<p:modifyVerifier/></p:presentation>'));}
  if(variant==='wrong-type')p.package.set('[Content_Types].xml',p.package.text('[Content_Types].xml').replaceAll('application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml','application/xml'));
  if(variant==='body-duplicate')p.package.set(notes,p.package.text(notes).replace('type="sldNum"','type="body"'));
  const before=p.package.toBytes();expect(()=>p.slides[0]!.inspectNotesText()).toThrow();expect(p.package.toBytes()).toEqual(before);
 }
});
test('serialization failure rolls back notes and retains a usable target',async()=>{
 const p=await deck(),s=p.slides[0]!,target=s.inspectNotesText(),before=p.package.toBytes(),original=p.package.toBytes.bind(p.package);p.package.toBytes=()=>{throw Error('serialize failed');};try{expect(()=>s.replaceNotesAt(target,'changed')).toThrow('serialize failed');}finally{p.package.toBytes=original;}expect(p.package.toBytes()).toEqual(before);s.replaceNotesAt(target,'retry');expect(s.inspectNotesText().text).toBe('retry');
});

test('eight canonical existing-notes scenarios execute saved and in-memory predicates',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{join}=await import('node:path'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings,scenarioIds}=await import('../acceptance/notes-editing.ts');
 const path='workflows/pptx/notes.feature',selected=selectSharedScenarios(path,await Bun.file(join(fixturesRoot(),path)).text(),scenarioIds),feature={...selected,scenarios:selected.scenarios.filter(s=>s.lifecycle==='implemented')},count=(n:number)=>({implemented:n,planned:0,total:n});
 const cases=feature.scenarios.flatMap(s=>s.cases),result=await executeAcceptance({root:'.',features:[feature],counts:{features:count(1),scenarios:count(8),cases:count(8),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}},bindings,'notes-edit-unit');expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(8);
});

test('notes body edits preserve non-body placeholders, UTF-16 encoding and namespace aliases',async()=>{
 const {utf16}=await import('../fixtures/admission.ts');const p=await deck();p.package.set(notes,utf16('<?xml version="1.0" encoding="UTF-16"?>'+p.package.text(notes).replace(/^<\?xml[^?]*\?>/,'').replaceAll('a:','d:').replace('xmlns:a=','xmlns:d=')));
 const bytes=p.package.toBytes(),fresh=await Presentation.open(bytes),s=fresh.slides[0]!,text=fresh.package.text(notes),doc=parseXml(text),P='http://schemas.openxmlformats.org/presentationml/2006/main';
 const unchanged=elements(doc,'sp',P).filter(n=>!elements(n,'ph',P).some(ph=>attribute(ph,'type')==='body')).map(n=>text.slice(n.start,n.end));
 s.replaceNotesAt(s.inspectNotesText(),'first\n second ');const saved=fresh.package.toBytes(),reopened=await Presentation.open(saved);expect(reopened.slides[0]!.inspectNotesText().text).toBe('first\n second ');expect([...reopened.package.get(notes)!.slice(0,2)]).toEqual([255,254]);for(const fragment of unchanged)expect(reopened.package.text(notes)).toContain(fragment);
});
test('unknown notes extensions and removed slide enrollment refuse rather than editing hidden contexts',async()=>{
 const p=await deck();p.package.set(notes,p.package.text(notes).replace('</p:notes>','<p:extLst/></p:notes>'));const before=p.package.toBytes();expect(()=>p.slides[0]!.inspectNotesText()).toThrow();expect(p.package.toBytes()).toEqual(before);
 const q=await deck(),main=q.package.mainPart();q.package.set(main,q.package.text(main).replace(/<p:sldId\b[^>]*\/>/,''));const removed=q.package.toBytes();expect(()=>q.slides[0]!.inspectNotesText()).toThrow();expect(q.package.toBytes()).toEqual(removed);
});

test('notes body properties and root structure reject unknown nested policy surfaces',async()=>{
 for(const [from,to]of [['<a:bodyPr/>','<a:bodyPr><a:scene3d/></a:bodyPr>'],['<a:lstStyle/>','<a:lstStyle><a:lvl1pPr/></a:lstStyle>'],['</p:notes>','<p:unknown/></p:notes>'],['<a:bodyPr/>','<a:bodyPr unknown="yes"/>']]){
  const p=await deck();p.package.set(notes,p.package.text(notes).replace(from!,to!));const bytes=p.package.toBytes();expect(()=>p.slides[0]!.inspectNotesText()).toThrow();expect(p.package.toBytes()).toEqual(bytes);
 }
});
