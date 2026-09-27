import {test,expect} from 'bun:test';import {mkdtemp,rm} from 'node:fs/promises';import {tmpdir} from 'node:os';import {join} from 'node:path';
import {OpcPackage} from '../../src/opc/index.ts';import {inspectComments,setCommentResolved,inspectCommentThreads,setCommentThreadResolved} from '../../src/docx/comments.ts';import {commentFixture} from '../acceptance/comments-docx.ts';
const EX='word/commentsExtended.xml',C='word/comments.xml',W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const states=(p:OpcPackage)=>inspectComments(p).comments.map(c=>[c.id,c.resolved,c.parentId??null]);

test('thread inspection returns two nonempty roots and exact detached replies without changing bytes',async()=>{
 const p=await commentFixture(),before=p.toBytes(),r=inspectCommentThreads(p);expect(r.unsupported).toEqual([]);expect(r.threads.map(t=>[t.root.id,t.replies.map(c=>c.id)])).toEqual([['0',[]],['1',['2']]]);expect(r.threads[1]!.root.text).toBe('Classical hubris');expect(r.threads[1]!.replies[0]!.text).toBe('(this is a threaded reply)');expect(r.threads[1]!.replies[0]!.parentId).toBe('1');for(const v of [r,r.threads,r.threads[0],r.threads[1]!.root,r.threads[1]!.replies,r.threads[1]!.replies[0],r.unsupported])expect(Object.isFrozen(v)).toBe(true);expect(p.toBytes()).toEqual(before);setCommentResolved(p,'1',true);expect(r.threads[1]!.root.resolved).toBe(false);expect(inspectCommentThreads(p).threads[1]!.root.resolved).toBe(true);
});

test('resolving a reply updates its complete existing thread while retaining unrelated root bodies anchors and members through path reopen',async()=>{
 const p=await commentFixture(),before=await OpcPackage.open(p.toBytes());expect(setCommentThreadResolved(p,'02',true)).toEqual({rootId:'1',commentIds:['1','2'],changed:2,changedParts:[EX]});expect(states(p)).toEqual([['0',false,null],['1',true,null],['2',true,'1']]);expect(p.text(C)).toBe(before.text(C));expect(p.text(p.mainPart())).toBe(before.text(before.mainPart()));expect(p.names()).toEqual(before.names());for(const n of before.names())if(n!==EX)expect(p.get(n)).toEqual(before.get(n));
 const dir=await mkdtemp(join(tmpdir(),'comment-thread-'));try{const file=join(dir,'resolved.docx');await p.save(file);const read=await OpcPackage.open(file);expect(states(read)).toEqual(states(p));expect(inspectCommentThreads(read).threads[1]!.replies).toHaveLength(1);expect(setCommentThreadResolved(read,'1',false)).toEqual({rootId:'1',commentIds:['1','2'],changed:2,changedParts:[EX]});await read.save(file);const reopened=await OpcPackage.open(file);for(const n of before.names())expect(reopened.get(n)).toEqual(before.get(n));}finally{await rm(dir,{recursive:true,force:true});}
});

test('mixed states only change mismatched flags and repeat requests preserve exact archives while single-comment semantics remain unchanged',async()=>{
 const p=await commentFixture();setCommentResolved(p,'1',true);expect(states(p)).toEqual([['0',false,null],['1',true,null],['2',false,'1']]);expect(setCommentThreadResolved(p,'2',true).changed).toBe(1);const before=p.toBytes();expect(setCommentThreadResolved(p,'1',true)).toEqual({rootId:'1',commentIds:['1','2'],changed:0,changedParts:[]});expect(p.toBytes()).toEqual(before);expect(setCommentThreadResolved(p,'0',true).commentIds).toEqual(['0']);expect(states(p)).toEqual([['0',true,null],['1',true,null],['2',true,'1']]);
});

async function nested(){const p=await commentFixture(),xml=p.text(C),match=xml.match(/<w:comment\b[^>]*w:id="2"[^]*?<\/w:comment>/)!;expect(match).not.toBeNull();const clone=match[0].replace('w:id="2"','w:id="3"').replaceAll('0E00AF3F','12345678').replace('(this is a threaded reply)','Nested reply');p.set(C,xml.replace('</w:comments>',clone+'</w:comments>'));p.set(EX,p.text(EX).replace('</w15:commentsEx>','<w15:commentEx w15:paraId="12345678" w15:paraIdParent="0E00AF3F" w15:done="0"/></w15:commentsEx>'));return p;}

test('nested descendants and out-of-order extension entries resolve to the root and retain parent IDs',async()=>{
 const p=await nested();const source=p.text(EX),entries=[...source.matchAll(/<w15:commentEx\b[^>]*\/>/g)].map(m=>m[0]);p.set(EX,source.replace(/<w15:commentEx\b[^>]*\/>/g,'').replace('</w15:commentsEx>',entries.reverse().join('')+'</w15:commentsEx>'));const before=p.text(C);expect(inspectCommentThreads(p).threads.map(t=>[t.root.id,t.replies.map(r=>[r.id,r.parentId])])).toEqual([['0',[]],['1',[['2','1'],['3','2']]]]);expect(setCommentThreadResolved(p,'3',true)).toEqual({rootId:'1',commentIds:['1','2','3'],changed:3,changedParts:[EX]});expect(states(p)).toEqual([['0',false,null],['1',true,null],['2',true,'1'],['3',true,'2']]);expect(p.text(C)).toBe(before);
});

test('missing root metadata refuses even if selected reply already has the requested state',async()=>{
 const p=await commentFixture();p.set(EX,p.text(EX).replace(/<w15:commentEx\b[^>]*w15:paraId="3395B541"[^>]*\/>/,''));const before=p.toBytes();expect(()=>setCommentThreadResolved(p,'2',false)).toThrow(expect.objectContaining({code:'docx-comments-missing-extension'}));expect(p.toBytes()).toEqual(before);expect(inspectCommentThreads(p).threads[1]!.root.id).toBe('1');
});

test('invalid arguments unsafe graphs unsupported bodies and protection refuse atomically including no-ops',async()=>{
 const p=await commentFixture(),before=p.toBytes();for(const [id,value]of [[null,true],['',true],['-1',true],['1.5',true],['9007199254740992',true],['99',true],['1','true']]){expect(()=>setCommentThreadResolved(p,id as string,value as boolean)).toThrow();expect(p.toBytes()).toEqual(before);}
 for(const kind of ['missing-extension','duplicate-id','duplicate-para','missing-parent','cycle','invalid-done','wrong-mime','external-link','protected','revision-body','orphan-entry']){const p=await commentFixture(kind),before=p.toBytes();expect(()=>setCommentThreadResolved(p,'2',false)).toThrow();expect(p.toBytes()).toEqual(before);}
});

test('thread inspection retains explicit unsupported findings but mutation does not silently ignore unsupported siblings',async()=>{
 const p=await commentFixture('revision-body'),before=p.toBytes(),r=inspectCommentThreads(p);expect(r.threads).toHaveLength(2);expect(r.unsupported.length).toBeGreaterThan(0);expect(Object.isFrozen(r.unsupported[0])).toBe(true);expect(()=>setCommentThreadResolved(p,'2',true)).toThrow(expect.objectContaining({code:'docx-comments-unsupported'}));expect(p.toBytes()).toEqual(before);
});

test('injected write or serialization failure restores all flags and retains prior package edits',async()=>{
 for(const stage of ['set','toBytes'] as const){const p=await nested();p.set(p.mainPart(),p.text(p.mainPart()).replace('</w:body>','<w:p><w:r><w:t>Prior</w:t></w:r></w:p></w:body>'));const before=p.toBytes(),original=p[stage].bind(p);let reached=false;if(stage==='set')p.set=(n,v)=>{original(n as never,v as never);reached=true;throw Error('injected write');};else p.toBytes=()=>{reached=true;throw Error('injected serialization');};try{expect(()=>setCommentThreadResolved(p,'3',true)).toThrow('injected');expect(reached).toBe(true);}finally{if(stage==='set')p.set=original as OpcPackage['set'];else p.toBytes=original as OpcPackage['toBytes'];}expect(p.toBytes()).toEqual(before);expect(states(p)).toEqual([['0',false,null],['1',false,null],['2',false,'1'],['3',false,'2']]);}
});

test('thread updates retain alias namespaces lexical values and UTF8 BOM or UTF16 byte order',async()=>{
 for(const encoding of ['bom','le','be']){const p=await nested();let x=p.text(EX).replaceAll('w15:','q:').replace('xmlns:w15=','xmlns:q=').replace('q:done="0"',"q:done = 'false'");if(encoding==='bom')p.set(EX,new Uint8Array([239,187,191,...new TextEncoder().encode(x)]));else{x=x.replace('UTF-8','UTF-16');const b=new Uint8Array(2+x.length*2);b.set(encoding==='le'?[255,254]:[254,255]);const v=new DataView(b.buffer);for(let i=0;i<x.length;i++)v.setUint16(2+2*i,x.charCodeAt(i),encoding==='le');p.set(EX,b);}const before=await OpcPackage.open(p.toBytes());expect(setCommentThreadResolved(p,'3',true).changed).toBe(3);const read=await OpcPackage.open(p.toBytes());expect(states(read).map(s=>s[1])).toEqual([false,true,true,true]);expect([...read.get(EX)!.slice(0,encoding==='bom'?3:2)]).toEqual(encoding==='bom'?[239,187,191]:encoding==='le'?[255,254]:[254,255]);expect(read.text(EX)).toContain("q:done = 'false'");for(const n of before.names())if(n!==EX)expect(read.get(n)).toEqual(before.get(n));}
});

test('single-comment edits also retain an existing UTF8 BOM and empty thread inspection returns explicit empty arrays',async()=>{
 const p=await commentFixture(),bytes=p.get(EX)!;p.set(EX,new Uint8Array([239,187,191,...bytes]));setCommentResolved(p,'1',true);expect([...p.get(EX)!.slice(0,3)]).toEqual([239,187,191]);const {Document}=await import('../../src/index.ts'),empty=await OpcPackage.open(Document.create().package.toBytes()),before=empty.toBytes();expect(inspectCommentThreads(empty)).toEqual({threads:[],unsupported:[]});expect(empty.toBytes()).toEqual(before);
});

test('comment source order including descendants before roots remains deterministic without changing thread ownership',async()=>{
 const p=await nested(),source=p.text(C),chunks=[...source.matchAll(/<w:comment\b[^]*?<\/w:comment>/g)].map(m=>m[0]);expect(chunks).toHaveLength(4);p.set(C,source.replace(/<w:comment\b[^]*?<\/w:comment>/g,'').replace('</w:comments>',[chunks[3],chunks[2],chunks[0],chunks[1]].join('')+'</w:comments>'));const before=p.toBytes(),r=inspectCommentThreads(p);expect(r.threads.map(t=>[t.root.id,t.replies.map(c=>c.id)])).toEqual([['0',[]],['1',['3','2']]]);expect(p.toBytes()).toEqual(before);expect(setCommentThreadResolved(p,'3',true).commentIds).toEqual(['3','2','1']);expect(inspectComments(p).comments.map(c=>[c.id,c.resolved])).toEqual([['3',true],['2',true],['0',false],['1',true]]);
});

test('batched flag edits preserve every unrelated character with quote entity and missing-flag variants',async()=>{
 const p=await nested(),before=p.text(EX);let source=before.replace('w15:done="0"','w15:done = \'off\'');source=source.replace('w15:paraId="3395B541" w15:done="0"','w15:paraId="3395B541" w15:done="&#48;"');source=source.replace('w15:paraId="12345678" w15:paraIdParent="0E00AF3F" w15:done="0"','w15:paraId="12345678" w15:paraIdParent="0E00AF3F"');p.set(EX,source);expect(setCommentThreadResolved(p,'2',true).changed).toBe(3);let expected=source.replace('w15:done="&#48;"','w15:done="1"').replace('w15:paraId="0E00AF3F" w15:paraIdParent="3395B541" w15:done="0"','w15:paraId="0E00AF3F" w15:paraIdParent="3395B541" w15:done="1"').replace('w15:paraId="12345678" w15:paraIdParent="0E00AF3F"/>','w15:paraId="12345678" w15:paraIdParent="0E00AF3F" w15:done="1"/>');expect(p.text(EX)).toBe(expected);expect(p.text(EX)).toContain("w15:done = 'off'");
});

test('thread operation count bound refuses without returning a truncated result or mutating flags',async()=>{
 const p=await commentFixture();p.set(C,`<w:comments xmlns:w="${W}">${Array.from({length:10001},(_,id)=>`<w:comment w:id="${id}"><w:p/></w:comment>`).join('')}</w:comments>`);p.delete(EX);const rel='word/_rels/document.xml.rels';p.set(rel,p.text(rel).replace(/<Relationship\b[^>]*Type="[^"]*\/commentsExtended"[^>]*\/>/,''));const before=p.toBytes();expect(()=>inspectCommentThreads(p)).toThrow(expect.objectContaining({code:'docx-comments-limit'}));expect(()=>setCommentThreadResolved(p,'0',false)).toThrow(expect.objectContaining({code:'docx-comments-limit'}));expect(p.toBytes()).toEqual(before);
});

test('existing-extension threads do not activate authored-comment response or root-only resolution profiles',async()=>{
 const {inventoryFeatures}=await import('../../scripts/gherkin.ts'),inv=await inventoryFeatures(process.cwd()),f=inv.features.find(f=>f.path.endsWith('/docx/comments.feature'))!,planned=f.scenarios.filter(s=>s.scenarioId.startsWith('@id-python-comments-'));expect(planned).toHaveLength(9);expect(planned.every(s=>s.lifecycle==='planned')).toBe(true);expect(inv.counts.cases.implemented).toBe(625);expect(inv.counts.cases.planned).toBe(46);
});
