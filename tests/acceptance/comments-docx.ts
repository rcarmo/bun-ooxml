import {expect} from 'bun:test';
import {fixturePath,F} from '../../scripts/fixture-inputs.ts';
import {OpcPackage,addPart,addRelationship} from '../../src/opc/index.ts';
import {inspectComments,setCommentResolved} from '../../src/docx/comments.ts';
import {OoxmlError} from '../../src/errors.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const fixture=fixturePath(F.officeWord.comments);
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export async function commentFixture(kind='plain'):Promise<OpcPackage>{
 const p=await OpcPackage.open(fixture),ex='word/commentsExtended.xml',comments='word/comments.xml',rels='word/_rels/document.xml.rels';
 if(kind==='missing-extension')p.set(rels,p.text(rels).replace(/<Relationship\b[^>]*Type="[^"]*\/commentsExtended"[^>]*\/>/,''));
 if(kind==='duplicate-id')p.set(comments,p.text(comments).replace('w:id="1"','w:id="00"'));
 if(kind==='duplicate-para')p.set(comments,p.text(comments).replace('w14:paraId="3395B541"','w14:paraId="103008CD"'));
 if(kind==='missing-parent')p.set(ex,p.text(ex).replace('w15:paraIdParent="3395B541"','w15:paraIdParent="ABCDEF12"'));
 if(kind==='cycle')p.set(ex,p.text(ex).replace('w15:paraId="3395B541"','w15:paraId="3395B541" w15:paraIdParent="0E00AF3F"'));
 if(kind==='invalid-done')p.set(ex,p.text(ex).replace('w15:done="0"','w15:done="maybe"'));
 if(kind==='wrong-mime')p.set('[Content_Types].xml',p.text('[Content_Types].xml').replace('application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml','application/vnd.ms-word.commentsExtended+xml'));
 if(kind==='external-link')p.set(rels,p.text(rels).replace('Target="commentsExtended.xml"','Target="https://example.invalid/comments.xml" TargetMode="External"'));
 if(kind==='protected'){const settings=p.related(p.mainPart(),'settings');if(settings)p.set(settings,p.text(settings).replace('</w:settings>','<w:documentProtection w:enforcement="1"/></w:settings>'));else{addPart(p,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');}}
 if(kind==='revision-body')p.set(comments,p.text(comments).replace('<w:r>','<w:ins w:id="77"><w:r>').replace('</w:r>','</w:r></w:ins>'));
 if(kind==='orphan-entry')p.set(ex,p.text(ex).replace('</w15:commentsEx>','<w15:commentEx w15:paraId="ABCDEF12" w15:done="0"/></w15:commentsEx>'));
 return p;
}
export const bindings:StepBinding[]=[
 {pattern:/^the pinned threaded Word comments package is opened$/,run:async c=>{c.pkg=await commentFixture();c.before=(c.pkg as OpcPackage).toBytes();}},
 {pattern:/^existing Word comments are inspected$/,run:c=>{c.inspection=inspectComments(c.pkg as OpcPackage);}},
 {pattern:/^the three comment bodies and reply parent match the pinned fixture$/,run:c=>{expect(c.inspection).toEqual({comments:[
  {id:'0',author:'Rui Carmo',initials:'RC',date:'2026-03-10T15:24:00Z',text:'This is a great opening',paragraphId:'103008CD',resolved:false},
  {id:'1',author:'Rui Carmo',initials:'RC',date:'2026-03-10T15:25:00Z',text:'Classical hubris',paragraphId:'3395B541',resolved:false},
  {id:'2',author:'Rui Carmo',initials:'RC',date:'2026-03-10T15:25:00Z',text:'(this is a threaded reply)',paragraphId:'0E00AF3F',parentId:'1',resolved:false},
 ],unsupported:[]});}},
 {pattern:/^the Word comment package bytes remain unchanged$/,run:c=>expect((c.pkg as OpcPackage).toBytes()).toEqual(c.before as Uint8Array)},
 {pattern:/^Word comment "([^"]+)" is marked resolved$/,run:(c,id)=>{c.receipt=setCommentResolved(c.pkg as OpcPackage,id,true);expect(c.receipt).toEqual({changedParts:['word/commentsExtended.xml'],changed:1});}},
 {pattern:/^saving and reopening shows that comment resolved and its reply link intact$/,run:async c=>{c.pkg=await OpcPackage.open((c.pkg as OpcPackage).toBytes());const v=inspectComments(c.pkg as OpcPackage);expect(v.comments.find(x=>x.id==='1')?.resolved).toBe(true);expect(v.comments.find(x=>x.id==='2')?.parentId).toBe('1');expect(v.comments.find(x=>x.id==='2')?.resolved).toBe(false);}},
 {pattern:/^only the existing commentsExtended part differs from the input$/,run:async c=>{const p=c.pkg as OpcPackage,b=await OpcPackage.open(c.before as Uint8Array);expect(p.names()).toEqual(b.names());for(const n of p.names())if(n==='word/commentsExtended.xml')expect(p.get(n)).not.toEqual(b.get(n));else expect(p.get(n)).toEqual(b.get(n));}},
 {pattern:/^Word comment "([^"]+)" is reopened$/,run:(c,id)=>{c.receipt=setCommentResolved(c.pkg as OpcPackage,id,false);}},
 {pattern:/^the original comment package member bytes are restored$/,run:async c=>{const p=await OpcPackage.open((c.pkg as OpcPackage).toBytes()),b=await OpcPackage.open(c.before as Uint8Array);expect(p.names()).toEqual(b.names());for(const n of b.names())expect(p.get(n)).toEqual(b.get(n));}},
 {pattern:/^an existing Word comment refusal package (\S+)$/,run:async(c,kind)=>{c.pkg=await commentFixture(kind);c.before=(c.pkg as OpcPackage).toBytes();}},
 {pattern:/^Word comment resolution is attempted$/,run:c=>{try{setCommentResolved(c.pkg as OpcPackage,'1',true);}catch(e){c.error=e;}}},
 {pattern:/^a typed comment refusal leaves every package byte unchanged$/,run:c=>{expect(c.error).toBeInstanceOf(OoxmlError);expect((c.pkg as OpcPackage).toBytes()).toEqual(c.before as Uint8Array);}},
];
