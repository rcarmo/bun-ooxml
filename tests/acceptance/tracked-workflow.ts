import {expect} from 'bun:test';
import {mkdtemp,rm,readdir} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {redlineFixture} from './redline-docx.ts';
import {OpcPackage,addRelationship} from '../../src/opc/index.ts';
import {inspectStories} from '../../src/docx/story.ts';
import {inspectRevisions,resolveRevisions} from '../../src/docx/revisions.ts';
import {patchOffice,type PatchRequest} from '../../src/workflow/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
const roots:string[]=[];
export async function cleanupTrackedWorkflowFixtures(){await Promise.all(roots.splice(0).map(p=>rm(p,{recursive:true,force:true})));}
export async function setupTrackedWorkflow(kind='plain'){
 const root=await mkdtemp(join(tmpdir(),'bun-tracked-batch-'));roots.push(root);const p=await redlineFixture(kind);
 if(kind==='external-settings')addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','https://example.invalid/settings.xml',{external:true});
 const source=join(root,'source.docx'),output=join(root,'output.docx'),before=p.toBytes();await Bun.write(source,before);await Bun.write(output,'output sentinel');
 return {root,source,output,before};
}
export const metadata={author:'Workflow reviewer',date:'2026-09-26T12:00:00Z'};
export function requestFor(state:Awaited<ReturnType<typeof setupTrackedWorkflow>>,variant:string):PatchRequest{
 const r:any={source:state.source,output:state.output,mode:variant==='preview'?'dry_run':'safe',changes:[{target:'within thirty',value:variant==='no-op'?'within thirty':variant==='deletion'?'':'within sixty'}],trackChanges:variant!=='untracked',revisionMetadata:metadata};
 if(variant==='missing')r.changes[0].target='absent';if(variant==='ambiguous')r.changes[0].target='thirty';
 if(variant==='multiple-targets')r.changes.push({target:'days.',value:'weeks.'});
 if(variant==='invalid-author')r.revisionMetadata={...metadata,author:' '};if(variant==='invalid-date')r.revisionMetadata={...metadata,date:'2026-02-30T12:00:00Z'};
 if(variant==='invalid-option')r.trackChanges='false';if(variant==='missing-metadata')delete r.revisionMetadata;if(variant==='non-docx')r.format='pptx';
 return r;
}
export const bindings:StepBinding[]=[
 {pattern:/^a tracked workflow package and an existing output sentinel$/,run:async c=>{Object.assign(c,await setupTrackedWorkflow());}},
 {pattern:/^a (tracked|untracked|no-op|preview|deletion) Word replacement batch is requested$/,run:async(c,op)=>{c.op=op;c.receipt=await patchOffice(requestFor(c as any,op));}},
 {pattern:/^the tracked workflow outcome is (committed|preview) with (\d+) committed changes and (\d+) committed revisions$/,run:async(c,outcome,changes,revisions)=>{
  const r=c.receipt as any;expect(r.status).toBe(outcome);expect(r.committedChanges).toBe(Number(changes));expect(r.trackedRevisions).toBe(Number(revisions));
  if(c.op==='preview'){expect(r.previewRevisions).toBe(2);expect(r.revisionIds).toEqual([]);expect(await Bun.file(c.output as string).text()).toBe('output sentinel');return;}
  expect(r.previewRevisions).toBe(0);const saved=await OpcPackage.open(c.output as string),observed=inspectRevisions(saved).revisions;expect(observed.length).toBe(Number(revisions));expect(r.revisionIds).toEqual(observed.map(x=>x.id));
  const expected=c.op==='no-op'?'Payment within thirty days.':c.op==='deletion'?'Payment  days.':'Payment within sixty days.';
  expect(inspectStories(saved).stories[0]!.paragraphs[0]!.text).toBe(expected);
  if(c.op==='tracked'||c.op==='deletion'){
   expect(observed.every(x=>x.author===metadata.author&&x.date===metadata.date)).toBe(true);
   for(const action of ['accept','reject'] as const){const copy=await OpcPackage.open(saved.toBytes());resolveRevisions(copy,action);expect(inspectStories(await OpcPackage.open(copy.toBytes())).stories[0]!.paragraphs[0]!.text).toBe(action==='accept'?expected:'Payment within thirty days.');}
  }
 }},
 {pattern:/^the tracked workflow source and unrelated members retain their bytes$/,run:async c=>{
  expect(await Bun.file(c.source as string).bytes()).toEqual(Uint8Array.from(c.before as Uint8Array));expect((await readdir(c.root as string)).sort()).toEqual(['output.docx','source.docx']);
  if(c.op==='preview')return;const before=await OpcPackage.open(c.before as Uint8Array),after=await OpcPackage.open(c.output as string);expect(after.names()).toEqual(before.names());for(const name of before.names())if(name!==before.mainPart())expect(after.get(name)).toEqual(before.get(name));
  if(c.op==='no-op')expect(await Bun.file(c.output as string).bytes()).toEqual(Uint8Array.from(c.before as Uint8Array));
 }},
 {pattern:/^a tracked workflow refusal package (\S+)$/,run:async(c,kind)=>{Object.assign(c,await setupTrackedWorkflow(kind));c.kind=kind;}},
 {pattern:/^a tracked Word batch is attempted$/,run:async c=>{c.receipt=await patchOffice(requestFor(c as any,c.kind as string));}},
 {pattern:/^the tracked workflow refuses with zero committed revisions and unchanged files$/,run:async c=>{
  const r=c.receipt as any;expect(r.status).toBe('refused');expect(r.committedChanges).toBe(0);expect(r.trackedRevisions).toBe(0);expect(r.previewRevisions).toBe(0);expect(r.revisionIds).toEqual([]);expect(r.changedParts).toEqual([]);
  expect(await Bun.file(c.source as string).bytes()).toEqual(Uint8Array.from(c.before as Uint8Array));expect(await Bun.file(c.output as string).text()).toBe('output sentinel');expect((await readdir(c.root as string)).sort()).toEqual(['output.docx','source.docx']);
 }},
];
