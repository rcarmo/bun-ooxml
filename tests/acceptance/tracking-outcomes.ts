import {expect} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {Document,OpcPackage} from '../../src/index.ts';
import {addPart,addRelationship,getContentType} from '../../src/opc/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import {inspectRevisions} from '../../src/docx/revisions.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings',MIME='application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml',PART='word/options/custom.xml';
const roots:string[]=[];
export async function cleanupTrackingOutcomes(){await Promise.all(roots.splice(0).map(p=>rm(p,{recursive:true,force:true})));}
const settings=(body:string)=>`<w:settings xmlns:w="${W}">${body}</w:settings>`;
const pkg=(d:Document)=>(d as unknown as {opcPackage:OpcPackage}).opcPackage;
type State={doc:Document;source:string;output:string;before:Uint8Array;parts:Map<string,Uint8Array>;held:Document['paragraphs'][number];reopened?:Document;xml?:string;layout?:string;receipt?:{changed:number;changedParts:string[]};errors:unknown[]};
const state=(c:Record<string,unknown>)=>c.tracking as State;
async function setup(c:Record<string,unknown>,configure?:(p:OpcPackage)=>void){
 const d=Document.create();d.addParagraph('Retained text');const p=await OpcPackage.open(d.package.toBytes());addPart(p,'customXml/opaque.bin',new Uint8Array([0,255,42]),'application/octet-stream');configure?.(p);
 const root=await mkdtemp(join(tmpdir(),'tracking-outcomes-'));roots.push(root);const source=join(root,'source.docx'),output=join(root,'output.docx'),before=p.toBytes();await Bun.write(source,before);const doc=await Document.open(source);
 c.tracking={doc,source,output,before,parts:new Map(doc.package.parts),held:doc.paragraphs[0]!,errors:[]} satisfies State;
}
function addSettings(p:OpcPackage,xml:string){addPart(p,PART,xml,MIME);addRelationship(p,p.mainPart(),REL,'options/custom.xml',{id:'settingsRef'});}
async function save(s:State){await s.doc.save(s.output);s.reopened=await Document.open(s.output);}
function capture(s:State,op:()=>unknown){try{op();s.errors.push(null);}catch(e){s.errors.push(e);}}
async function sourceUnchanged(s:State){expect(await Bun.file(s.source).bytes()).toEqual(Uint8Array.from(s.before));}
function sameMembers(s:State,after:Document,except:string[]=[]){expect([...after.package.parts.keys()].sort()).toEqual([...s.parts.keys()].sort());for(const[n,b]of s.parts)if(!except.includes(n))expect(after.package.get(n)).toEqual(b);}
export const bindings:StepBinding[]=[
 {pattern:/^a Word tracking package with text Retained text and an unrelated opaque payload$/,run:async c=>setup(c)},
 {pattern:/^the settings editor saves tracking (true|false) with author Session reviewer and reopens the output$/,run:async(c,on)=>{const s=state(c);s.doc.enableTracking('Session reviewer');if(on==='false')s.doc.disableTracking();await save(s);}},
 {pattern:/^the reopened tracking preference is (true|false) and the session author is empty$/,run:(c,on)=>{const d=state(c).reopened!;expect(d.trackChangesEnabled).toBe(on==='true');expect(d.trackAuthor).toBe('');}},
 {pattern:/^the output has one internal settings relationship with the Word settings content type$/,run:c=>{const p=pkg(state(c).reopened!),rels=p.relationships(p.mainPart()).filter(r=>r.type===REL);expect(rels).toHaveLength(1);expect(rels[0]!.external).toBe(false);expect(getContentType(p,rels[0]!.resolved!)).toBe(MIME);const flags=elements(parseXml(p.text(rels[0]!.resolved!)),'trackRevisions',W);expect(flags).toHaveLength(1);expect(attribute(flags[0]!,'val',W)).toBe(state(c).reopened!.trackChangesEnabled?'1':'0');}},
 {pattern:/^the reopened text and opaque payload equal the source and the source archive is unchanged$/,run:async c=>{const s=state(c);expect(s.reopened!.paragraphs.map(p=>p.text)).toEqual(['Retained text']);for(const[n,b]of s.parts)if(!['[Content_Types].xml','word/_rels/document.xml.rels'].includes(n))expect(s.reopened!.package.get(n)).toEqual(b);await sourceUnchanged(s);}},
 {pattern:/^a Word tracking package with (custom-prefix|UTF-16|UTF-8-BOM) settings and text Retained text$/,run:async(c,layout)=>{
  let xml=`<q:settings xmlns:q="${W}" xmlns:w="urn:foreign"><q:zoom q:percent='90'/><q:revisionView q:markup='0'/><q:doNotTrackMoves/></q:settings>`;
  if(layout==='UTF-16')xml='<?xml version="1.0" encoding="UTF-16"?>'+xml;
  await setup(c,p=>{addSettings(p,xml);if(layout==='UTF-16'){const bytes=new Uint8Array(2+xml.length*2);bytes.set([255,254]);const view=new DataView(bytes.buffer);for(let i=0;i<xml.length;i++)view.setUint16(2+i*2,xml.charCodeAt(i),true);p.set(PART,bytes);}else if(layout==='UTF-8-BOM')p.set(PART,new Uint8Array([239,187,191,...new TextEncoder().encode(xml)]));});Object.assign(state(c),{xml,layout});
 }},
 {pattern:/^the settings editor enables tracking with author Session reviewer and saves the output$/,run:async c=>{const s=state(c);s.doc.enableTracking('Session reviewer');await save(s);}},
 {pattern:/^the reopened preference is true and the settings path and encoding match the source$/,run:c=>{const s=state(c),d=s.reopened!,p=pkg(d);expect(d.trackChangesEnabled).toBe(true);expect(p.relationships(p.mainPart()).filter(r=>r.type===REL).map(r=>r.resolved)).toEqual([PART]);const length=s.layout==='UTF-16'?2:s.layout==='UTF-8-BOM'?3:1;expect(d.package.get(PART)!.slice(0,length)).toEqual(s.parts.get(PART)!.slice(0,length));}},
 {pattern:/^removing only the added tracking element reproduces the original settings XML$/,run:c=>{const s=state(c),xml=pkg(s.reopened!).text(PART),tree=parseXml(xml),flags=elements(tree,'trackRevisions',W);expect(flags).toHaveLength(1);const f=flags[0]!;expect(xml.slice(0,f.start)+xml.slice(f.end)).toBe(s.xml!);expect(tree.root.children.map(n=>n.localName)).toEqual(['zoom','revisionView','trackRevisions','doNotTrackMoves']);}},
 {pattern:/^every unrelated member and the source archive is unchanged$/,run:async c=>{const s=state(c);sameMembers(s,s.reopened!,[PART]);await sourceUnchanged(s);}},
 {pattern:/^a Word tracking package with (absent|on|off|true) preference$/,run:async(c,value)=>setup(c,p=>{if(value!=='absent')addSettings(p,settings(`<w:trackRevisions w:val='${value}'/>`));})},
 {pattern:/^the settings editor requests tracking (true|false) and changes the session author to Session reviewer$/,run:(c,on)=>{const s=state(c);s.receipt=on==='true'?s.doc.enableTracking('Session reviewer'):s.doc.disableTracking();s.doc.setTrackAuthor('Session reviewer');}},
 {pattern:/^the preference is (true|false) and the session author is Session reviewer$/,run:(c,on)=>{const d=state(c).doc;expect(d.trackChangesEnabled).toBe(on==='true');expect(d.trackAuthor).toBe('Session reviewer');}},
 {pattern:/^the complete archive is unchanged with zero changed parts$/,run:async c=>{const s=state(c);expect(s.receipt).toEqual({changed:0,changedParts:[]});expect(await s.doc.save()).toEqual(s.before);await sourceUnchanged(s);}},
 {pattern:/^a Word tracking package with (protected|external|duplicate-link|orphan|wrong-content-type|duplicate-flag|wrong-order|shared-owner) settings and session author Previous$/,run:async(c,defect)=>{
  await setup(c,p=>{
   if(defect==='external'){addRelationship(p,p.mainPart(),REL,'https://example.invalid/settings',{external:true});return;}
   addSettings(p,settings(defect==='protected'?'<w:documentProtection w:enforcement="1"/>':defect==='duplicate-flag'?'<w:trackRevisions/><w:trackRevisions/>':defect==='wrong-order'?'<w:compat/><w:trackRevisions/>':''));
   if(defect==='duplicate-link')addRelationship(p,p.mainPart(),REL,'options/custom.xml',{id:'secondSettings'});
   if(defect==='orphan'){const rel='word/_rels/document.xml.rels';p.set(rel,p.text(rel).replace(/<Relationship[^>]*Id="settingsRef"[^>]*\/>/,''));}
   if(defect==='wrong-content-type')p.set('[Content_Types].xml',p.text('[Content_Types].xml').replace(MIME,'application/xml'));
   if(defect==='shared-owner')addRelationship(p,'',REL,'word/options/custom.xml');
  });state(c).doc.setTrackAuthor('Previous');
 }},
 {pattern:/^the settings editor attempts enabling and disabling tracking$/,run:c=>{const s=state(c);capture(s,()=>s.doc.enableTracking('Next'));expect(s.doc.package.toBytes()).toEqual(s.before);expect(s.doc.trackAuthor).toBe('Previous');capture(s,()=>s.doc.disableTracking());}},
 {pattern:/^both requests refuse and the session author remains Previous$/,run:c=>{const s=state(c);expect(s.errors).toHaveLength(2);for(const e of s.errors)expect(e).toBeInstanceOf(Error);expect(s.doc.trackAuthor).toBe('Previous');}},
 {pattern:/^the archive and held paragraph text Retained text are unchanged$/,run:async c=>{const s=state(c);expect(await s.doc.save()).toEqual(s.before);sameMembers(s,s.doc);expect(s.held.text).toBe('Retained text');await sourceUnchanged(s);}},
 {pattern:/^a Word tracking package without settings and with session author Previous$/,run:async c=>{await setup(c);state(c).doc.setTrackAuthor('Previous');}},
 {pattern:/^enabling tracking fails at (part-write|relationship-write|serialization)$/,run:(c,stage)=>{const s=state(c),p=pkg(s.doc),set=p.set.bind(p),serialize=p.toBytes.bind(p);if(stage==='serialization')p.toBytes=()=>{throw Error('injected serialization');};else p.set=(n,v)=>{set(n,v);if(n===(stage==='part-write'?'word/settings.xml':'word/_rels/document.xml.rels'))throw Error('injected '+stage);};try{capture(s,()=>s.doc.enableTracking('Next'));}finally{p.set=set;p.toBytes=serialize;}expect(String(s.errors[0])).toContain('injected');}},
 {pattern:/^the settings operation refuses and the session author remains Previous$/,run:c=>{const s=state(c);expect(s.errors).toHaveLength(1);expect(s.errors[0]).toBeInstanceOf(Error);expect(s.doc.trackAuthor).toBe('Previous');}},
 {pattern:/^tracking is enabled and an ordinary text edit writes Explicit plain edit then saves and reopens$/,run:async c=>{const s=state(c);s.doc.enableTracking('Session reviewer');s.doc.paragraphs[0]!.setText('Explicit plain edit');await save(s);}},
 {pattern:/^the reopened text is (.+) and the tracking preference is true$/,run:(c,text)=>{const d=state(c).reopened!;expect(d.paragraphs.map(p=>p.text)).toEqual([text!]);expect(d.trackChangesEnabled).toBe(true);}},
 {pattern:/^the reopened document contains no revisions and the opaque payload is unchanged$/,run:async c=>{const s=state(c);expect(inspectRevisions(pkg(s.reopened!))).toEqual({revisions:[],unsupported:[]});expect(s.reopened!.package.get('customXml/opaque.bin')).toEqual(s.parts.get('customXml/opaque.bin')!);await sourceUnchanged(s);}},
 {pattern:/^the settings editor attempts to enable tracking with (blank|XML-control|non-string) author input$/,run:(c,kind)=>{const s=state(c),author=kind==='blank'?'   ':kind==='XML-control'?'\u0000':1;capture(s,()=>s.doc.enableTracking(author as string));}},
];
