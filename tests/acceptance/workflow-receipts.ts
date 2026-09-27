import assert from 'node:assert/strict';
import {mkdtemp,readdir,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {loadMutationFixtures,verifyFixture} from '../../scripts/shared-fixtures.ts';
import {fixturePath} from '../../scripts/fixture-inputs.ts';
import {Document,Presentation} from '../../src/index.ts';
import {findPlaceholderText} from '../../src/pptx/placeholders.ts';
import {patchOffice,type PatchReceipt} from '../../src/workflow/index.ts';

type State={root:string;source:string;output:string;sourceBytes:Uint8Array;outputBytes:Uint8Array;entries:string[];format:'docx'|'pptx';original:string;changes:{target:string;value:string}[];receipt?:PatchReceipt};
const state=(c:Record<string,unknown>)=>c.state as State;
const roots=new Set<string>();
const sha=(bytes:Uint8Array)=>new Bun.CryptoHasher('sha256').update(bytes).digest('hex');
export async function cleanupWorkflowReceipts(){const pending=[...roots];await Promise.all(pending.map(async root=>{await rm(root,{recursive:true,force:true});roots.delete(root);}));}
async function setup(c:Record<string,unknown>,id:string,format:State['format']){
 const fixture=(await loadMutationFixtures()).fixtures.find(f=>f.id===id);assert(fixture);
 const bytes=await Bun.file(fixturePath(fixture.assetId)).bytes();await verifyFixture(bytes,fixture);
 const root=await mkdtemp(join(tmpdir(),'bun-ooxml-receipts-'));roots.add(root);
 const s=state(c);s.root=root;s.source=join(root,id);s.output=join(root,'existing-output.'+format);s.format=format;
 await Bun.write(s.source,bytes);await Bun.write(s.output,'existing destination sentinel');
 s.sourceBytes=await Bun.file(s.source).bytes();s.outputBytes=await Bun.file(s.output).bytes();s.entries=(await readdir(root)).sort();
}
async function unchanged(s:State){
 assert.deepEqual(await Bun.file(s.source).bytes(),s.sourceBytes);
 assert.deepEqual(await Bun.file(s.output).bytes(),s.outputBytes);
 assert.deepEqual((await readdir(s.root)).sort(),s.entries);
 if(s.format==='pptx'){const deck=await Presentation.open(s.source);assert.equal(findPlaceholderText(deck.slides[0]!,'title').text,s.original);}
 else{const doc=await Document.open(s.source);assert.equal(doc.paragraphs.map(p=>p.text).join('\n'),s.original);}
}
function receipt(s:State,status:PatchReceipt['status']){
 const r=s.receipt;assert(r);assert.equal(r.status,status);assert.equal(r.sourceSha256,sha(s.sourceBytes));assert.equal(r.outputSha256,undefined);
 assert.deepEqual(r.results.map(({target,value})=>({target,value})),s.changes);return r;
}
export const bindings:StepBinding[]=[
 {pattern:/^a presentation with the title "([^"]+)"$/,run:async(c,title)=>{await setup(c,'title-and-subtitle.pptx','pptx');const s=state(c),deck=await Presentation.open(s.source);assert.equal(findPlaceholderText(deck.slides[0]!,'title').text,title);s.original=title;}},
 {pattern:/^a title change to "([^"]+)" is previewed$/,run:async(c,value)=>{const s=state(c);s.changes=[{target:'slide:1/title',value}];s.receipt=await patchOffice({source:s.source,output:s.output,mode:'dry_run',changes:s.changes});}},
 {pattern:/^the preview identifies the original target and requested replacement$/,run:c=>{const s=state(c),r=receipt(s,'preview');assert.equal(r.error,undefined);assert.equal(r.results.length,1);assert.equal(r.results[0]!.target,'slide:1/title');assert.equal(r.results[0]!.value,s.changes[0]!.value);assert.equal(r.results[0]!.matched,1);assert.equal(r.results[0]!.status,'matched');assert.equal(r.results[0]!.code,undefined);}},
 {pattern:/^the preview reports zero committed changes$/,run:async c=>{const s=state(c),r=receipt(s,'preview');assert.equal(r.committedChanges,0);await unchanged(s);}},
 {pattern:/^a document containing "([^"]+)" once and not containing "([^"]+)"$/,run:async(c,present,missing)=>{await setup(c,'present-placeholder.docx','docx');const s=state(c),d=await Document.open(s.source);assert.equal(d.find(present).length,1);assert.equal(d.find(missing).length,0);s.original=d.paragraphs.map(p=>p.text).join('\n');s.changes=[{target:present,value:'Changed present'},{target:missing,value:'Must not apply'}];}},
 {pattern:/^both placeholders are resolved before mutation$/,run:async c=>{const s=state(c);s.receipt=await patchOffice({source:s.source,output:s.output,mode:'strict',changes:s.changes});}},
 {pattern:/^"([^"]+)" reports exactly (one|zero) match(?:es)?$/,run:async(c,target,count)=>{const s=state(c),r=receipt(s,'refused');assert.equal(r.error?.code,'workflow-target-missing');assert.equal(r.committedChanges,0);assert.deepEqual(r.changedParts,[]);const matches=r.results.filter(v=>v.target===target);assert.equal(matches.length,1);const result=matches[0]!;assert.equal(result.matched,count==='one'?1:0);assert.equal(result.status,count==='one'?'matched':'unmatched');assert.equal(result.code,count==='one'?undefined:'workflow-target-missing');await unchanged(s);}},
];
