import {test,expect} from 'bun:test';
import {mkdtemp,mkdir,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {inventoryFeatures,executeAcceptance,selectSharedScenarios} from '../../scripts/gherkin.ts';
import {runAcceptance} from '../../scripts/acceptance.ts';
const path='references/fixtures-ooxml/workflows/mixed.feature';
const text='@planned\nFeature: Mixed runtimes\n @profile-common @id-bound\n Scenario: Implemented outcome\n  Given a bound input\n  Then a bound outcome\n @profile-other @id-other\n Scenario: Other policy\n  Given an unsupported input\n  Then an unsupported outcome\n';
async function project(config:any){const root=await mkdtemp(join(tmpdir(),'bun-shared-select-'));await mkdir(join(root,'features'),{recursive:true});await mkdir(join(root,'references/fixtures-ooxml/workflows'),{recursive:true});await Bun.write(join(root,path),text);await Bun.write(join(root,'features/shared.json'),JSON.stringify(config));return root;}
const entry={path,lifecycle:'implemented',runner:'bun',scenarioIds:['@id-bound']};
test('shared scenario selection retains unsupported outcomes as planned without invoking their steps',async()=>{
 const root=await project({schemaVersion:2,features:[entry]});try{
  const inventory=await inventoryFeatures(root);expect(inventory.counts.scenarios).toEqual({implemented:1,planned:1,total:2});expect(inventory.counts.cases).toEqual({implemented:1,planned:1,total:2});
  const feature=inventory.features[0]!;expect(feature.sourceSha256).toBe(new Bun.CryptoHasher('sha256').update(text).digest('hex'));
  let calls=0;const result=await executeAcceptance(inventory,[{pattern:/^a bound (input|outcome)$/,run:(context)=>{calls++;expect((context.scenario as {lifecycle:string}).lifecycle).toBe('implemented');expect((context.case as {lifecycle:string}).lifecycle).toBe('implemented');expect((context.feature as {tags:string[]}).tags).toEqual(['@planned']);}}],'selection-test');
  expect(calls).toBe(2);expect(result.failures).toEqual([]);expect(result.counts.features).toEqual({passed:1,failed:0,planned:0,total:1});expect(result.counts.cases).toEqual({passed:1,failed:0,planned:1,total:2});
  expect(result.features[0]!.scenarios[1]!.cases[0]!.steps.every(s=>s.status==='planned')).toBe(true);
 }finally{await rm(root,{recursive:true,force:true});}
});
test('selection rejects implicit whole-file implementation, unknown IDs, duplicates and empty selections',async()=>{
 expect(()=>selectSharedScenarios(path,text,['@id-missing'])).toThrow('Unknown');
 expect(()=>selectSharedScenarios(path,text,['@id-bound','@id-bound'])).toThrow('Duplicate');
 expect(()=>selectSharedScenarios(path,text,[])).toThrow('empty');
 expect(()=>selectSharedScenarios(path,text.replace('@planned','@implemented @bun'),['@id-bound'])).toThrow('planned');
 for(const config of [{schemaVersion:1,features:[entry]},{schemaVersion:2,features:[{path,lifecycle:'implemented',runner:'bun'}]}]){
  const root=await project(config);try{await expect(inventoryFeatures(root)).rejects.toThrow('shared');}finally{await rm(root,{recursive:true,force:true});}
 }
});
test('full acceptance rejects planned scenarios inside an otherwise implemented feature',async()=>{
 const root=await project({schemaVersion:2,features:[entry]});try{
  await expect(runAcceptance([{pattern:/^a bound (input|outcome)$/,run:()=>{}}],{root,full:true})).rejects.toThrow('planned');
  const report=await Bun.file(join(root,'artifacts/acceptance.json')).json();expect(report.execution.cases.passed).toBe(1);expect(report.execution.cases.planned).toBe(1);
 }finally{await rm(root,{recursive:true,force:true});}
});

test('outline rows with shared display names execute and report every distinct identity',async()=>{
 const outlines=text+' @id-outline\n Scenario Outline: Shared display name\n  Given input <value>\n  Then output <value>\n  Examples:\n   | value |\n   | one |\n   | two |\n';
 const f=selectSharedScenarios(path,outlines,['@id-bound']);const rows=f.scenarios.find(s=>s.scenarioId==='@id-outline')!.cases;
 expect(rows).toHaveLength(2);expect(new Set(rows.map(c=>c.identityKey)).size).toBe(2);
 const selected=selectSharedScenarios(path,outlines,['@id-outline']);
 const count=(n:number)=>({implemented:n,planned:0,total:n});const seen:string[]=[];
 const result=await executeAcceptance({root:'.',features:[selected],counts:{features:count(1),scenarios:count(1),cases:count(2),steps:count(4)}},[
  {pattern:/^input (one|two)$/,run:(_c,value)=>{seen.push(value!);}},
  {pattern:/^output (one|two)$/,run:(_c,value)=>{if(value==='two')throw Error('second row predicate');}},
 ],'same-display');
 expect(seen).toEqual(['one','two']);expect(result.counts.cases.passed).toBe(1);expect(result.counts.cases.failed).toBe(1);expect(result.counts.cases.planned).toBe(2);
 const reports=result.features[0]!.scenarios.find(s=>s.scenarioId==='@id-outline')!.cases;
 expect(reports.map(c=>c.identityKey)).toEqual(rows.map(c=>c.identityKey));expect(new Set(reports.map(c=>c.caseId)).size).toBe(2);
 expect(()=>selectSharedScenarios(path,outlines.replace('| two |','| one |'),['@id-outline'])).toThrow('Duplicate expanded case identity');
});
