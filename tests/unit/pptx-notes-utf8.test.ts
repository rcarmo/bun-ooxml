import {test,expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Presentation,OpcPackage} from '../../src/index.ts';
import {fixturePath,F} from '../../scripts/fixture-inputs.ts';
const part='ppt/notesSlides/notesSlide1.xml',utf8=(s:string)=>new TextEncoder().encode(s);
const source=()=>Bun.file(fixturePath(F.goSlides.notes)).bytes();
async function deck(){return Presentation.open(await source());}

test('UTF8 notes replacement owns byte input, saves exact Unicode text and preserves every other member',async()=>{
 const bytes=await source(),d=await Presentation.open(bytes),s=d.slides[0]!,target=s.inspectNotesText(),value=' Leading 雪 & 😀\nlast ',input=utf8(value),before=input.slice(),parts=new Map(d.package.names().map(n=>[n,d.package.get(n)!]));
 expect(s.replaceNotesUtf8At(target,input)).toEqual({changedParts:[part]});expect(input).toEqual(before);input.fill(0);expect(s.inspectNotesText().text).toBe(value);expect(()=>s.replaceNotesUtf8At(target,utf8('stale'))).toThrow(expect.objectContaining({code:'PPTX_NOTES_STALE'}));
 const root=await mkdtemp(join(tmpdir(),'notes-utf8-'));try{const path=join(root,'notes.pptx');await d.save(path);const after=await Presentation.open(path);expect(after.slides[0]!.inspectNotesText().text).toBe(value);expect(after.package.names()).toEqual([...parts.keys()].sort());for(const[n,b]of parts)if(n!==part)expect(after.package.get(n)).toEqual(b);}finally{await rm(root,{recursive:true,force:true});}
});

test('malformed UTF8 sequences refuse as encoding errors without consuming the held target',async()=>{
 const d=await deck(),s=d.slides[0]!,t=s.inspectNotesText(),before=d.package.toBytes();
 for(const raw of [[255],[128],[192,175],[226,130],[237,160,128],[244,144,128,128],[239,187,191,255]]){const input=new Uint8Array(raw),copy=input.slice();expect(()=>s.replaceNotesUtf8At(t,input)).toThrow(expect.objectContaining({code:'PPTX_NOTES_UTF8_INVALID'}));expect(input).toEqual(copy);expect(d.package.toBytes()).toEqual(before);}
 expect(s.replaceNotesUtf8At(t,utf8(t.text))).toEqual({changedParts:[]});expect(d.package.toBytes()).toEqual(before);expect(s.replaceNotesUtf8At(t,utf8('retry'))).toEqual({changedParts:[part]});
});

test('valid UTF8 XML controls and line policy still refuse through the existing notes editor',async()=>{
 const d=await deck(),s=d.slides[0]!,t=s.inspectNotesText(),before=d.package.toBytes();for(const text of ['bad\0','two\tcolumns','a\rb','a\r\nb','\u0001','\ufffe']){expect(()=>s.replaceNotesUtf8At(t,utf8(text))).toThrow();expect(d.package.toBytes()).toEqual(before);}expect(s.replaceNotesUtf8At(t,utf8(t.text))).toEqual({changedParts:[]});expect(d.package.toBytes()).toEqual(before);
});

test('UTF8 byte no-op retains archive and handles; empty and leading BOM bytes are literal text',async()=>{
 const d=await deck(),s=d.slides[0]!,t=s.inspectNotesText(),before=d.package.toBytes();expect(s.replaceNotesUtf8At(t,utf8(t.text))).toEqual({changedParts:[]});expect(d.package.toBytes()).toEqual(before);
 const withBom='\ufeff'+t.text;expect(s.replaceNotesUtf8At(t,utf8(withBom))).toEqual({changedParts:[part]});expect(s.inspectNotesText().text).toBe(withBom);expect((await Presentation.open(d.package.toBytes())).slides[0]!.inspectNotesText().text).toBe(withBom);
 const current=s.inspectNotesText();s.replaceNotesUtf8At(current,new Uint8Array());expect(s.inspectNotesText().text).toBe('');const empty=s.inspectNotesText(),saved=d.package.toBytes();expect(s.replaceNotesUtf8At(empty,new Uint8Array())).toEqual({changedParts:[]});expect(d.package.toBytes()).toEqual(saved);s.replaceNotesUtf8At(empty,utf8('refilled'));expect(s.inspectNotesText().text).toBe('refilled');
});

test('foreign forged wrong-slide and externally changed notes targets refuse bytes including no-op input',async()=>{
 const d=await deck(),other=await deck(),s=d.slides[0]!,t=s.inspectNotesText(),before=d.package.toBytes(),ob=other.package.toBytes();expect(()=>other.slides[0]!.replaceNotesUtf8At(t,utf8('foreign'))).toThrow(expect.objectContaining({code:'PPTX_NOTES_STALE'}));expect(other.package.toBytes()).toEqual(ob);expect(()=>s.replaceNotesUtf8At({...t},utf8(t.text))).toThrow(expect.objectContaining({code:'PPTX_NOTES_STALE'}));expect(()=>d.slides[1]!.replaceNotesUtf8At(t,utf8('foreign'))).toThrow(expect.objectContaining({code:'PPTX_NOTES_STALE'}));expect(d.package.toBytes()).toEqual(before);
 for(const path of [part,'ppt/slides/_rels/slide1.xml.rels']){const fresh=await deck(),slide=fresh.slides[0]!,anchor=slide.inspectNotesText();fresh.package.set(path,fresh.package.text(path)+' ');const changed=fresh.package.toBytes();expect(()=>slide.replaceNotesUtf8At(anchor,utf8(anchor.text))).toThrow(expect.objectContaining({code:'PPTX_NOTES_STALE'}));expect(fresh.package.toBytes()).toEqual(changed);}
});

test('byte admission rejects shared detached spoofed and oversized views before caller hooks',async()=>{
 const d=await deck(),s=d.slides[0]!,t=s.inspectNotesText(),before=d.package.toBytes();let called=false;
 const large=new Uint8Array(8*1024*1024+1);Object.defineProperty(large,'byteLength',{get(){called=true;return 1;}});expect(()=>s.replaceNotesUtf8At(t,large)).toThrow(expect.objectContaining({code:'PPTX_NOTES_BYTE_LIMIT'}));expect(called).toBe(false);
 const shared=new Uint8Array(new SharedArrayBuffer(1));shared[0]=65;Object.defineProperty(shared,'buffer',{get(){called=true;return new ArrayBuffer(1);}});expect(()=>s.replaceNotesUtf8At(t,shared)).toThrow(expect.objectContaining({code:'PPTX_NOTES_BYTE_SOURCE'}));expect(called).toBe(false);
 const detached=utf8('abc');structuredClone(detached,{transfer:[detached.buffer]});for(const input of [detached,null,'text',new ArrayBuffer(4),new DataView(new ArrayBuffer(4)),[65],Object.create(Uint8Array.prototype),new Proxy(utf8('abc'),{})])expect(()=>s.replaceNotesUtf8At(t,input as Uint8Array)).toThrow(expect.objectContaining({code:'PPTX_NOTES_BYTE_SOURCE'}));expect(d.package.toBytes()).toEqual(before);expect(s.replaceNotesUtf8At(t,utf8(t.text))).toEqual({changedParts:[]});
});

test('Buffer subview decoding ignores iterator slice and buffer getters and preserves adjacent bytes',async()=>{
 const d=await deck(),s=d.slides[0]!,t=s.inspectNotesText(),input=Buffer.from('XX雪😀YY'),view=input.subarray(2,input.length-2),before=Uint8Array.from(input);let called=false;
 for(const field of ['buffer','byteLength','byteOffset'])Object.defineProperty(view,field,{get(){called=true;throw Error('getter');}});Object.defineProperty(view,'slice',{value:()=>{throw Error('slice');}});Object.defineProperty(view,Symbol.iterator,{value:()=>{throw Error('iterator');}});
 expect(s.replaceNotesUtf8At(t,view)).toEqual({changedParts:[part]});expect(s.inspectNotesText().text).toBe('雪😀');expect(called).toBe(false);expect(Uint8Array.from(input)).toEqual(before);
});

test('package write and serialization faults roll back byte replacements and permit same-target retry',async()=>{
 for(const stage of ['set','toBytes'] as const){const d=await deck(),s=d.slides[0]!,t=s.inspectNotesText(),before=d.package.toBytes(),pkg=d.package,original=pkg[stage].bind(pkg);if(stage==='set')pkg.set=(n,v)=>{(original as OpcPackage['set'])(n,v);throw Error('injected write');};else pkg.toBytes=()=>{throw Error('injected serialization');};try{expect(()=>s.replaceNotesUtf8At(t,utf8('changed'))).toThrow('injected');}finally{if(stage==='set')pkg.set=original as OpcPackage['set'];else pkg.toBytes=original as OpcPackage['toBytes'];}expect(d.package.toBytes()).toEqual(before);expect(s.replaceNotesUtf8At(t,utf8('retry'))).toEqual({changedParts:[part]});expect(s.inspectNotesText().text).toBe('retry');}
});

test('byte replacement retains UTF16 notes-part codec and non-body placeholders after reopen',async()=>{
 const {utf16}=await import('../fixtures/admission.ts'),d=await deck(),text=d.package.text(part),notes=utf16('<?xml version="1.0" encoding="UTF-16"?>'+text.replace(/^<\?xml[^?]*\?>/,''));d.package.set(part,notes);const source=d.package.toBytes(),loaded=await Presentation.open(source),s=loaded.slides[0]!,t=s.inspectNotesText();s.replaceNotesUtf8At(t,utf8('雪😀'));
 const after=await Presentation.open(loaded.package.toBytes());expect(after.slides[0]!.inspectNotesText().text).toBe('雪😀');expect([...after.package.get(part)!.slice(0,2)]).toEqual([255,254]);expect(after.package.text(part)).toBe(new TextDecoder('utf-16le').decode(notes).replace(t.text,'雪😀'));for(const n of loaded.package.names())if(n!==part)expect(after.package.get(n)).toEqual(d.package.get(n));
});

test('canonical foreign invalid-byte and no-op case executes all four refusals and exact archive custody',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),path='workflows/native/pptx-text.feature',text=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const selected=selectSharedScenarios(path,text,['@id-pptx-go-notes-refusal-and-noop-custody']),feature={...selected,scenarios:selected.scenarios.filter(s=>s.lifecycle==='implemented')};
 const result=await executeAcceptance({root:'.',features:[feature],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(7)}},bindings,'notes-utf8-canonical');expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(1);expect(result.counts.steps.passed).toBe(7);
});

test('byte entry retains output-size and slide-enrollment refusal with unchanged package bytes',async()=>{
 const d=await deck(),s=d.slides[0]!,t=s.inspectNotesText(),before=d.package.toBytes();expect(()=>s.replaceNotesUtf8At(t,new Uint8Array(8*1024*1024).fill(65))).toThrow();expect(d.package.toBytes()).toEqual(before);expect(s.replaceNotesUtf8At(t,utf8(t.text))).toEqual({changedParts:[]});
 const main=d.package.mainPart();d.package.set(main,d.package.text(main).replace(/<p:sldId\b[^>]*\/>/,''));const changed=d.package.toBytes();expect(()=>s.replaceNotesUtf8At(t,utf8('changed'))).toThrow(expect.objectContaining({code:'PPTX_STALE_SLIDE'}));expect(d.package.toBytes()).toEqual(changed);
});

test('late protection and unsupported notes topology refuse byte no-ops without destructive repair',async()=>{
 for(const variant of ['protection','field','lock']){const d=await deck(),s=d.slides[0]!,t=s.inspectNotesText();if(variant==='protection'){const main=d.package.mainPart();d.package.set(main,d.package.text(main).replace('</p:presentation>','<p:modifyVerifier/></p:presentation>'));}else d.package.set(part,variant==='field'?d.package.text(part).replace('<a:r><a:t>Remember','<a:r><a:fld/> <a:t>Remember'):d.package.text(part).replaceAll('<a:spLocks noGrp="1"/>','<a:spLocks noTextEdit="1"/>'));const before=d.package.toBytes();expect(()=>s.replaceNotesUtf8At(t,utf8(t.text))).toThrow();expect(d.package.toBytes()).toEqual(before);}
});

test('compound notes assertions detect each missing refusal, wrong decoding error and changed archive',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),path='workflows/native/pptx-text.feature',text=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const selected=selectSharedScenarios(path,text,['@id-pptx-go-notes-refusal-and-noop-custody']),feature={...selected,scenarios:selected.scenarios.filter(s=>s.lifecycle==='implemented')},inventory={root:'.',features:[feature],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(7)}};
 type State={refusals:unknown[];original:Uint8Array;deck:Presentation;foreignDeck:Presentation};
 const step='the first session attempts JSON notes text "bad\\u0000", one invalid UTF-8 byte FF, and JSON text "two\\tcolumns"';
 for(const mutate of [...[0,1,2,3].map(index=>(s:State)=>{s.refusals[index]=undefined;}),(s:State)=>{s.refusals=[];},(s:State)=>{s.refusals[2]=Error('wrong decoder');},(s:State)=>{s.deck.package.set(part,s.deck.package.text(part)+' ');},(s:State)=>{s.foreignDeck.package.set(part,s.foreignDeck.package.text(part)+' ');},(s:State)=>{s.original[0]=0;}]){
  expect(bindings.filter(b=>b.pattern.test(step))).toHaveLength(1);const corrupt=bindings.map(b=>b.pattern.test(step)?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);mutate(c.state as State);}}:b);const result=await executeAcceptance(inventory,corrupt,'notes-refusal-control');expect(result.counts.cases.failed).toBe(1);expect(result.counts.steps.failed).toBe(1);expect(result.counts.steps.undefined).toBe(0);expect(result.counts.steps.ambiguous).toBe(0);
 }
 const noOp='replacing through the held target with its unchanged text succeeds';const corrupt=bindings.map(b=>b.pattern.test(noOp)?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);const s=c.state as State;s.deck.package.set(part,s.deck.package.text(part)+' ');}}:b);const result=await executeAcceptance(inventory,corrupt,'notes-noop-control');expect(result.counts.cases.failed).toBe(1);expect(result.counts.steps.failed).toBe(1);expect(result.counts.steps.undefined).toBe(0);expect(result.counts.steps.ambiguous).toBe(0);
});
