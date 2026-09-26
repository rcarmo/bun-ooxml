import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {inspectComments,setCommentResolved} from '../../src/docx/comments.ts';
import {OpcPackage} from '../../src/opc/package.ts';
import {bindings,commentFixture} from '../acceptance/comments-docx.ts';

test('all existing comment Gherkin cases assert saved outcomes',async()=>{
 const path='references/fixtures-ooxml/workflows/docx/comments.feature',f=parseFeature(path,(await Bun.file(join(fixturesRoot(), "workflows/docx/comments.feature")).text()).replace(/^@planned/m, '@implemented @bun'));const cases=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inventory:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(cases.length),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}};
 const result=await executeAcceptance(inventory,bindings,newAcceptanceRunId());expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(14);
});

test('renamed extension prefixes and multiline comments preserve exact unrelated bytes',async()=>{
 const p=await commentFixture(),ex='word/commentsExtended.xml';p.set(ex,p.text(ex).replaceAll('xmlns:w15=','xmlns:thread=').replaceAll('w15:','thread:'));
 const before=p.text(ex);setCommentResolved(p,'1',true);expect(p.text(ex)).toBe(before.replace('thread:paraId="3395B541" thread:done="0"','thread:paraId="3395B541" thread:done="1"'));
 const reopened=await OpcPackage.open(p.toBytes());expect(inspectComments(reopened).comments[1]!.resolved).toBe(true);
});

test('unsupported comment body is reported read-only and no resolution mutates it',async()=>{
 const p=await commentFixture('revision-body'),before=p.toBytes();expect(inspectComments(p).unsupported.length).toBeGreaterThan(0);expect(p.toBytes()).toEqual(before);expect(()=>setCommentResolved(p,'1',true)).toThrow();expect(p.toBytes()).toEqual(before);
});

test('unknown IDs and runtime nonboolean arguments leave package unchanged',async()=>{
 const p=await commentFixture(),before=p.toBytes();for(const id of ['missing','-1','999'])expect(()=>setCommentResolved(p,id,true)).toThrow();
 expect(()=>setCommentResolved(p,'1','false' as unknown as boolean)).toThrow();expect(p.toBytes()).toEqual(before);
});

test('attribute edits preserve quote style, spacing, entities and unrelated attribute text',async()=>{
 const p=await commentFixture(),part='word/commentsExtended.xml';
 const xml=p.text(part).replace('w15:paraId="3395B541" w15:done="0"',`w15:paraId="3395B541" note="literal w15:done='0'" w15:done = '&#48;'`);
 p.set(part,xml);const before=p.toBytes();
 expect(setCommentResolved(p,'1',true)).toEqual({changed:1,changedParts:[part]});
 expect(p.text(part)).toBe(xml.replace("w15:done = '&#48;'","w15:done = '1'"));
 const reopened=await OpcPackage.open(p.toBytes());expect(inspectComments(reopened).comments.find(c=>c.id==='1')?.resolved).toBe(true);
 const original=await OpcPackage.open(before);for(const name of p.names())if(name!==part)expect(p.get(name)).toEqual(original.get(name));
});
test('missing done flag uses an attribute prefix under a default element namespace',async()=>{
 const p=await commentFixture(),part='word/commentsExtended.xml';
 const xml=p.text(part).replaceAll('w15:commentsEx','commentsEx').replaceAll('w15:commentEx','commentEx').replace('xmlns:w15=','xmlns=').replaceAll('w15:','q:').replace('<commentsEx ', '<commentsEx xmlns:q="http://schemas.microsoft.com/office/word/2012/wordml" ').replace('q:paraId="3395B541" q:done="0"','q:paraId="3395B541"');
 p.set(part,xml);const before=p.toBytes();expect(setCommentResolved(p,'1',false).changed).toBe(0);expect(p.toBytes()).toEqual(before);
 expect(setCommentResolved(p,'1',true).changed).toBe(1);expect(p.text(part)).toBe(xml.replace('q:paraId="3395B541"/>','q:paraId="3395B541" q:done="1"/>'));
 expect(inspectComments(await OpcPackage.open(p.toBytes())).comments.find(c=>c.id==='1')?.resolved).toBe(true);
});
test('same boolean state retains on/off spelling and exact archive bytes',async()=>{
 for(const [word,state]of [['on',true],['true',true],['off',false],['false',false]] as const){
  const p=await commentFixture(),part='word/commentsExtended.xml';p.set(part,p.text(part).replace('w15:paraId="3395B541" w15:done="0"',`w15:paraId="3395B541" w15:done="${word}"`));const before=p.toBytes();
  expect(setCommentResolved(p,'1',state)).toEqual({changed:0,changedParts:[]});expect(p.toBytes()).toEqual(before);
 }
});
test('serialization failure rolls back the done flag while retaining earlier edits',async()=>{
 const p=await commentFixture(),part='word/commentsExtended.xml';
 p.set('word/document.xml',p.text('word/document.xml').replace('</w:body>','<!--prior edit--></w:body>'));const before=p.toBytes(),xml=p.text(part);
 const serialize=p.toBytes;p.toBytes=()=>{throw new Error('injected serialization failure');};
 try{expect(()=>setCommentResolved(p,'1',true)).toThrow('injected serialization failure');expect(p.text(part)).toBe(xml);}finally{p.toBytes=serialize;}
 expect(p.toBytes()).toEqual(before);
});
test('misqualified done or parent metadata refuses instead of being ignored',async()=>{
 for(const key of ['done','paraIdParent']){
  const p=await commentFixture(),part='word/commentsExtended.xml';p.set(part,p.text(part).replace(`w15:${key}=`,`bogus:${key}=`).replace('<w15:commentsEx ','<w15:commentsEx xmlns:bogus="urn:wrong" '));const before=p.toBytes();
  expect(()=>setCommentResolved(p,'1',true)).toThrow();expect(p.toBytes()).toEqual(before);
 }
});

test('resolution refuses every enforcing settings part and accepts only explicit disabled protection',async()=>{
 const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
 for(const enforcement of [undefined,'','1','true','on','unexpected','false','0','OFF']){
  const p=await commentFixture(),settings=p.related(p.mainPart(),'settings')!;
  expect(settings).toBeDefined();
  p.set(settings,`<s:settings xmlns:s="${W}"><s:documentProtection${enforcement===undefined?'':` s:enforcement="${enforcement}"`}/></s:settings>`);
  const before=p.toBytes();
  if(['false','0','OFF'].includes(enforcement??'')){expect(setCommentResolved(p,'1',true).changed).toBe(1);}else{expect(()=>setCommentResolved(p,'1',true)).toThrow();expect(p.toBytes()).toEqual(before);}
 }
 const p=await commentFixture(),settings=p.related(p.mainPart(),'settings')!;
 p.set(settings,`<s:settings xmlns:s="${W}"><s:documentProtection s:enforcement="0"/><s:documentProtection s:enforcement="1"/></s:settings>`);const before=p.toBytes();expect(()=>setCommentResolved(p,'1',true)).toThrow();expect(p.toBytes()).toEqual(before);
});
test('inspection results are detached and cannot bypass package custody',async()=>{
 const p=await commentFixture(),before=p.toBytes(),first=inspectComments(p);first.comments[1]!.resolved=true;first.comments[2]!.parentId='999';first.comments[0]!.text='mutated';first.unsupported.push({part:'word/comments.xml',kind:'invented'});
 const second=inspectComments(p);expect(second.comments[1]!.resolved).toBe(false);expect(second.comments[2]!.parentId).toBe('1');expect(second.comments[0]!.text).toBe('This is a great opening');expect(second.unsupported).toEqual([]);expect(p.toBytes()).toEqual(before);
});
test('comment paragraph newlines and tabs are read without rewriting bodies',async()=>{
 const p=await commentFixture(),part='word/comments.xml';const xml=p.text(part).replace('This is a great opening','This is a great opening</w:t><w:tab/><w:t>tab</w:t><w:br/><w:t>line');p.set(part,xml);const before=p.toBytes();
 expect(inspectComments(p).comments[0]!.text).toBe('This is a great opening\ttab\nline');expect(p.toBytes()).toEqual(before);
 setCommentResolved(p,'1',true);expect(p.text(part)).toBe(xml);expect((await OpcPackage.open(p.toBytes())).text(part)).toBe(xml);
});
test('non-self-closing entries retain boundary whitespace exactly',async()=>{
 const p=await commentFixture(),part='word/commentsExtended.xml';const xml=p.text(part).replace('w15:paraId="3395B541" w15:done="0"/>','w15:paraId="3395B541" \n > \r\n </w15:commentEx>');p.set(part,xml);
 setCommentResolved(p,'1',true);expect(p.text(part)).toBe(xml.replace('w15:paraId="3395B541" \n >','w15:paraId="3395B541" w15:done="1" \n >'));expect(inspectComments(p).comments[1]!.resolved).toBe(true);
});
test('multiple comments settings and external settings refuse before mutation',async()=>{
 const {addPart,addRelationship}=await import('../../src/opc/index.ts');const p=await commentFixture(),W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
 addPart(p,'word/extra-settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="true"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),R,'extra-settings.xml');let before=p.toBytes();expect(()=>setCommentResolved(p,'1',true)).toThrow();expect(p.toBytes()).toEqual(before);
 const q=await commentFixture();addRelationship(q,q.mainPart(),R,'https://example.invalid/settings.xml',{external:true});before=q.toBytes();expect(()=>setCommentResolved(q,'1',true)).toThrow();expect(q.toBytes()).toEqual(before);
});

test('UTF-16 extension encoding and BOM survive a done-only edit',async()=>{
 const p=await commentFixture(),part='word/commentsExtended.xml',xml=p.text(part).replace('encoding="UTF-8"','encoding="UTF-16"');
 const encode=(value:string)=>{const out=new Uint8Array(2+value.length*2),view=new DataView(out.buffer);view.setUint16(0,0xfeff,true);for(let i=0;i<value.length;i++)view.setUint16(2+i*2,value.charCodeAt(i),true);return out;};
 p.set(part,encode(xml));setCommentResolved(p,'1',true);expect(p.get(part)).toEqual(encode(xml.replace('w15:paraId="3395B541" w15:done="0"','w15:paraId="3395B541" w15:done="1"')));
 expect(inspectComments(await OpcPackage.open(p.toBytes())).comments[1]!.resolved).toBe(true);
});
test('misqualified comment IDs and paragraph metadata refuse the whole resolution',async()=>{
 for(const [from,to]of [['w:id="1"','id="1"'],['w14:paraId="3395B541"','paraId="3395B541"']] as const){
  const p=await commentFixture(),part='word/comments.xml';p.set(part,p.text(part).replace(from,to));const before=p.toBytes();expect(()=>setCommentResolved(p,'1',true)).toThrow();expect(p.toBytes()).toEqual(before);
 }
});
test('incomplete extended metadata remains inspectable but absent targets cannot be resolved',async()=>{
 const p=await commentFixture(),part='word/commentsExtended.xml';p.set(part,p.text(part).replace(/<w15:commentEx w15:paraId="103008CD"[^>]*\/>/,''));
 const before=p.toBytes();expect(inspectComments(p).comments).toHaveLength(3);expect(()=>setCommentResolved(p,'0',true)).toThrow();expect(p.toBytes()).toEqual(before);
 expect(setCommentResolved(p,'1',true).changed).toBe(1);
});

test('public entrypoints save and reopen the resolved flag without changing comment bodies',async()=>{
 const {mkdtemp,rm}=await import('node:fs/promises'),{tmpdir}=await import('node:os');
 const api=await import('../../src/index.ts'),docx=await import('../../src/docx/index.ts');
 expect(api.inspectComments).toBe(docx.inspectComments);expect(api.setCommentResolved).toBe(docx.setCommentResolved);
 const root=await mkdtemp(join(tmpdir(),'bun-comment-save-'));
 try{const p=await commentFixture(),comments=p.get('word/comments.xml')!,path=join(root,'resolved.docx');api.setCommentResolved(p,'1',true);await p.save(path);const q=await OpcPackage.open(path);expect(api.inspectComments(q).comments[1]!.resolved).toBe(true);expect(q.get('word/comments.xml')).toEqual(comments);}finally{await rm(root,{recursive:true,force:true});}
});
test('a document without comment parts remains unchanged when inspected or refused',async()=>{
 const {Document}=await import('../../src/docx/index.ts'),p=await OpcPackage.open(await Document.create().save()),before=p.toBytes();
 expect(inspectComments(p)).toEqual({comments:[],unsupported:[]});expect(()=>setCommentResolved(p,'0',true)).toThrow();expect(p.toBytes()).toEqual(before);
});
