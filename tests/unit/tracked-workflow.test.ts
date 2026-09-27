import{afterEach,test,expect}from'bun:test';
import{join}from'node:path';
import{fixturesRoot}from'../../scripts/fixture-inputs.ts';
import{parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory}from'../../scripts/gherkin.ts';
import{bindings,cleanupTrackedWorkflowFixtures,setupTrackedWorkflow,requestFor}from'../acceptance/tracked-workflow.ts';
import{patchOffice}from'../../src/workflow/index.ts';
afterEach(cleanupTrackedWorkflowFixtures);
test('tracked Word workflow dispatch executes17outcome/refusal cases',async()=>{
 const {sharedScenarios}=await import('../helpers/shared-scenarios.ts');const [f]=await sharedScenarios(['@id-docx-track-changes-option-outcome','@id-docx-workflow-tracked-refusal']);if(!f)throw Error('Missing tracked workflow');const cases=f.scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const inventory:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(cases.length),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}};
 const result=await executeAcceptance(inventory,bindings,newAcceptanceRunId());expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(17);
});
test('initial destination precondition refusal reports zero revision counts',async()=>{
 const s=await setupTrackedWorkflow(),r=await patchOffice({...requestFor(s,'tracked'),expectedDestinationSha256:null});expect(r.status).toBe('refused');expect((r as any).trackedRevisions).toBe(0);expect((r as any).revisionIds).toEqual([]);expect(await Bun.file(s.output).text()).toBe('output sentinel');
});

test('output serialization/save refusal clears staged revision identities',async()=>{
 const {OpcPackage}=await import('../../src/opc/package.ts'),s=await setupTrackedWorkflow();const save=OpcPackage.prototype.save;
 OpcPackage.prototype.save=async()=>{throw new Error('injected final save failure');};
 try{const r=await patchOffice(requestFor(s,'tracked'));expect(r.status).toBe('refused');expect(r.trackedRevisions).toBe(0);expect(r.previewRevisions).toBe(0);expect(r.revisionIds).toEqual([]);expect(r.changedParts).toEqual([]);expect(await Bun.file(s.output).text()).toBe('output sentinel');expect(await Bun.file(s.source).bytes()).toEqual(Uint8Array.from(s.before));}finally{OpcPackage.prototype.save=save;}
});
test('strict tracked overwrite and explicit false dispatch have distinct saved outcomes',async()=>{
 const {OpcPackage}=await import('../../src/opc/package.ts'),{inspectRevisions}=await import('../../src/docx/revisions.ts'),{inspectStories}=await import('../../src/docx/story.ts');
 const s=await setupTrackedWorkflow(),request=requestFor(s,'tracked');delete request.output;request.mode='strict';
 const r=await patchOffice(request);expect(r.status).toBe('committed');expect(r.trackedRevisions).toBe(2);const p=await OpcPackage.open(s.source);expect(inspectRevisions(p).revisions.map(x=>x.id)).toEqual(r.revisionIds);expect(inspectStories(p,{view:'original'}).stories[0]!.paragraphs[0]!.text).toBe('Payment within thirty days.');expect(await Bun.file(s.output).text()).toBe('output sentinel');
 const t=await setupTrackedWorkflow(),untracked=requestFor(t,'untracked');untracked.revisionMetadata={author:'',date:'invalid'};const plain=await patchOffice(untracked);expect(plain.status).toBe('committed');expect(plain.trackedRevisions).toBe(0);expect(inspectRevisions(await OpcPackage.open(t.output)).revisions).toEqual([]);
});
test('tracked no-op still validates metadata and refuses unsupported topology',async()=>{
 for(const variant of ['protected','existing','external-settings','drawing']){const s=await setupTrackedWorkflow(variant),request=requestFor(s,'no-op');const r=await patchOffice(request);expect(r.status).toBe('refused');expect(r.trackedRevisions).toBe(0);expect(await Bun.file(s.output).text()).toBe('output sentinel');}
 const s=await setupTrackedWorkflow(),request=requestFor(s,'no-op');request.revisionMetadata={author:'Reviewer',date:'2026-02-30T12:00:00Z'};expect((await patchOffice(request)).status).toBe('refused');
});
test('overlapping literal occurrences never become one tracked target',async()=>{
 const {OpcPackage}=await import('../../src/opc/package.ts'),s=await setupTrackedWorkflow(),p=await OpcPackage.open(s.source);p.set(p.mainPart(),p.text(p.mainPart()).replace('thirty days.','aaa'));await p.save(s.source);const before=await Bun.file(s.source).bytes();const request=requestFor(s,'tracked');request.changes=[{target:'aa',value:'x'}];const r=await patchOffice(request);expect(r.status).toBe('refused');expect(r.results[0]?.matched).toBe(2);expect(r.trackedRevisions).toBe(0);expect(await Bun.file(s.source).bytes()).toEqual(before);expect(await Bun.file(s.output).text()).toBe('output sentinel');
});

test('malformed settings root and runtime metadata do not bypass tracked preflight',async()=>{
 const {OpcPackage,addPart,addRelationship}=await import('../../src/opc/index.ts');
 const s=await setupTrackedWorkflow(),p=await OpcPackage.open(s.source);addPart(p,'word/settings.xml','<wrong/>','application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');addRelationship(p,p.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');await p.save(s.source);const before=await Bun.file(s.source).bytes();
 const r=await patchOffice(requestFor(s,'tracked'));expect(r.status).toBe('refused');expect(r.trackedRevisions).toBe(0);expect(await Bun.file(s.source).bytes()).toEqual(before);expect(await Bun.file(s.output).text()).toBe('output sentinel');
 const t=await setupTrackedWorkflow();for(const revisionMetadata of [null,{}, {author:3,date:'2026-09-26T12:00:00Z'},{author:'A',date:false}]){const request=requestFor(t,'tracked') as any;request.revisionMetadata=revisionMetadata;const x=await patchOffice(request);expect(x.status).toBe('refused');expect(x.trackedRevisions).toBe(0);expect(x.previewRevisions).toBe(0);}
});

test('caller mutation after dispatch cannot change the requested tracking mode or revision metadata',async()=>{
 const {OpcPackage}=await import('../../src/opc/package.ts'),{inspectRevisions}=await import('../../src/docx/revisions.ts');
 const s=await setupTrackedWorkflow(),request=requestFor(s,'tracked');request.revisionMetadata={author:'Original reviewer',date:'2026-09-26T12:00:00Z'};
 const pending=patchOffice(request);request.trackChanges=false;request.revisionMetadata.author='Mutated reviewer';request.changes[0]!.value='unrequested value';
 const result=await pending;expect(result.status).toBe('committed');expect(result.trackedRevisions).toBe(2);expect(result.results[0]?.value).toBe('within sixty');
 const revisions=inspectRevisions(await OpcPackage.open(s.output)).revisions;expect(revisions.every(r=>r.author==='Original reviewer')).toBe(true);expect(revisions).toHaveLength(2);
});

test('destination change after staging refuses without publishing staged revisions',async()=>{
 const {OpcPackage}=await import('../../src/opc/package.ts'),{writeFileSync}=await import('node:fs'),s=await setupTrackedWorkflow();
 const serialize=OpcPackage.prototype.toBytes;let staged=false;
 OpcPackage.prototype.toBytes=function(){const bytes=serialize.call(this);if(!staged&&this.text(this.mainPart()).includes('Workflow reviewer')){staged=true;writeFileSync(s.output,'external writer');}return bytes;};
 try{const r=await patchOffice(requestFor(s,'tracked'));expect(staged).toBe(true);expect(r.status).toBe('refused');expect(r.error?.code).toBe('workflow-stale-destination');expect(r.trackedRevisions).toBe(0);expect(r.revisionIds).toEqual([]);expect(r.changedParts).toEqual([]);expect(await Bun.file(s.output).text()).toBe('external writer');expect(await Bun.file(s.source).bytes()).toEqual(Uint8Array.from(s.before));}finally{OpcPackage.prototype.toBytes=serialize;}
});
