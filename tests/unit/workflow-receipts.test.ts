import {test,expect} from 'bun:test';
import {sharedScenarios} from '../helpers/shared-scenarios.ts';
import {selectSharedScenarios,executeAcceptance,type StepBinding} from '../../scripts/gherkin.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {bindings,cleanup} from '../acceptance/steps.ts';
import {mkdtemp,rm,lstat} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
import {runAcceptance} from '../../scripts/acceptance.ts';
import type {PatchReceipt} from '../../src/workflow/index.ts';
type State={root:string;source:string;output:string;receipt:PatchReceipt};
const path='workflows/workflow-receipts.feature';
const ids=['@id-office-preview-details','@id-office-docx-exact-match-counts'];
const paths=['workflows/pptx/mutation-safety.feature','workflows/docx/mutation-safety.feature'];
const sourceText=async()=>Object.fromEntries(await Promise.all(paths.map(async p=>[p,await Bun.file(join(fixturesRoot(),p)).text()])));
const replace=(source:Record<string,string>,from:string,to:string)=>Object.fromEntries(Object.entries(source).map(([p,s])=>[p,s.replace(from,to)]));
async function execute(source:Record<string,string>,custom:StepBinding[]=bindings){
 const features=paths.map(p=>{const selected=[ids[p.includes('/pptx/')?0:1]!],f=selectSharedScenarios(p,source[p]!,selected);return {...f,scenarios:f.scenarios.filter(s=>selected.includes(s.scenarioId))};});
 const count=(n:number)=>({implemented:n,planned:0,total:n});
 try{return await executeAcceptance({root:process.cwd(),features,counts:{features:count(2),scenarios:count(2),cases:count(2),steps:count(8)}},custom,'workflow-receipts-unit');}finally{await cleanup();}
}

test('retained receipt scenarios execute both exact Then predicates with no undefined bindings',async()=>{
 const result=await execute(await sourceText());expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(2);expect(result.counts.steps.passed).toBe(8);expect(result.counts.steps.undefined).toBe(0);expect(result.counts.steps.ambiguous).toBe(0);
});

const preview='a title change to "Changed by dry run" is previewed',counts='both placeholders are resolved before mutation';
function changed(step:string,mutate:(s:State)=>unknown|Promise<unknown>){
 expect(bindings.filter(b=>b.pattern.test(step))).toHaveLength(1);
 return bindings.map(b=>b.pattern.test(step)?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);await mutate(c.state as State);}}:b);
}
function failed(result:Awaited<ReturnType<typeof execute>>){expect(result.counts.cases.failed).toBe(1);expect(result.counts.cases.passed).toBe(1);expect(result.counts.steps.failed).toBe(1);expect(result.counts.steps.undefined).toBe(0);expect(result.counts.steps.ambiguous).toBe(0);}

test('preview Then assertions reject incorrect target value count status and fabricated commit claims',async()=>{
 const source=await sourceText();
 for(const mutate of [
  (s:State)=>{s.receipt.results[0]!.target='slide:1/subtitle';},
  (s:State)=>{s.receipt.results[0]!.value='Wrong';},
  (s:State)=>{s.receipt.results[0]!.matched=0;},
  (s:State)=>{s.receipt.results[0]!.status='committed';},
  (s:State)=>{s.receipt.results.push({...s.receipt.results[0]!});},
  (s:State)=>{s.receipt.results=[];},
  (s:State)=>{s.receipt.status='committed';},
  (s:State)=>{s.receipt.committedChanges=1;},
  (s:State)=>{s.receipt.outputSha256='fake';},
  (s:State)=>{s.receipt.sourceSha256='fake';},
  (s:State)=>{s.receipt.error={code:'fake',message:'fake'};},
 ])failed(await execute(source,changed(preview,mutate)));
});

test('placeholder Then assertions reject swapped counts missing duplicated or mislabelled results and false success',async()=>{
 const source=await sourceText();
 for(const mutate of [
  (s:State)=>{s.receipt.results[0]!.matched=0;},
  (s:State)=>{s.receipt.results[1]!.matched=1;},
  (s:State)=>{s.receipt.results[0]!.matched=2;},
  (s:State)=>{s.receipt.results.reverse();},
  (s:State)=>{s.receipt.results[1]={...s.receipt.results[0]!};},
  (s:State)=>{s.receipt.results.pop();},
  (s:State)=>{s.receipt.results[1]!.target='<Wrong>';},
  (s:State)=>{s.receipt.results[0]!.value='Wrong';},
  (s:State)=>{s.receipt.results[0]!.status='committed';},
  (s:State)=>{s.receipt.results[1]!.code='wrong-code';},
  (s:State)=>{s.receipt.error=undefined;},
  (s:State)=>{s.receipt.status='committed';s.receipt.committedChanges=1;},
  (s:State)=>{s.receipt.changedParts=['word/document.xml'];},
 ])failed(await execute(source,changed(counts,mutate)));
});

test('receipt predicates reject changed source destination and unexpected directory entries',async()=>{
 const source=await sourceText();
 for(const step of [preview,counts])for(const field of ['source','output','extra'] as const){
  failed(await execute(source,changed(step,async s=>{await Bun.write(field==='extra'?join(s.root,'unexpected.tmp'):s[field],'changed bytes');})));
 }
});

test('exact-count expected text and original-title fixture controls fail real assertions',async()=>{
 const source=await sourceText();
 for(const [from,to]of [['"<Present>" reports exactly one match','"<Present>" reports exactly zero matches'],['"<Missing>" reports exactly zero matches','"<Missing>" reports exactly one match'],['the title "Original title"','the title "Wrong title"']]){expect(Object.values(source).join('\n')).toContain(from!);failed(await execute(replace(source,from!,to!)));}
});

test('receipt temporary paths are removed after success failed Then and failed Given assertions',async()=>{
 const roots:string[]=[],source=await sourceText();
 const capture=bindings.map(b=>b.pattern.test('a presentation with the title "Original title"')||b.pattern.test('a document containing "<Present>" once and not containing "<Missing>"')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{try{await b.run(c,...args);}finally{roots.push((c.state as State).root);}}}:b);
 expect((await execute(source,capture)).failures).toEqual([]);
 failed(await execute(replace(source,'"<Missing>" reports exactly zero matches','"<Missing>" reports exactly one match'),capture));
 failed(await execute(replace(source,'the title "Original title"','the title "Wrong title"'),capture));
 expect(roots).toHaveLength(6);expect(new Set(roots).size).toBe(6);for(const root of roots)await expect(lstat(root)).rejects.toMatchObject({code:'ENOENT'});
});

test('acceptance runner cleans receipt fixtures on a failing binding without caller cleanup',async()=>{
 const root=await mkdtemp(join(tmpdir(),'receipt-runner-')),log=join(root,'paths.json'),steps=join(root,'steps.ts'),base=join(import.meta.dir,'../acceptance/steps.ts');
 try{
  const sources=await sourceText();for(const p of paths)await Bun.write(join(root,'references/fixtures-ooxml',p),sources[p]!);
  await Bun.write(join(root,'features/shared.json'),JSON.stringify({schemaVersion:2,features:paths.map((p,i)=>({path:'references/fixtures-ooxml/'+p,lifecycle:'implemented',runner:'bun',scenarioIds:[ids[i]]}))}));
  await Bun.write(steps,`import {bindings as base,cleanup} from ${JSON.stringify(base)}; import {writeFile} from 'node:fs/promises'; export {cleanup}; const paths:string[]=[]; export const bindings=base.map(b=>b.pattern.test(${JSON.stringify(preview)})||b.pattern.test(${JSON.stringify(counts)})?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args); paths.push((c.state as {root:string}).root); await writeFile(${JSON.stringify(log)},JSON.stringify(paths)); throw Error('injected after setup');}}:b);`);
  await expect(runAcceptance(undefined,{root,stepsModulePath:steps})).rejects.toThrow('injected after setup');
  const savedPaths=await Bun.file(log).json() as string[];expect(savedPaths).toHaveLength(2);for(const p of savedPaths)await expect(lstat(p)).rejects.toMatchObject({code:'ENOENT'});
  const result=await Bun.file(join(root,'artifacts/acceptance.json')).json();expect(result.status).toBe('failed');expect(result.execution.cases.failed).toBe(2);expect(result.execution.steps.undefined).toBe(0);
 }finally{await rm(root,{recursive:true,force:true});}
});
