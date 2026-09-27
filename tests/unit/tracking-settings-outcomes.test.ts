import {test,expect} from 'bun:test';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
import {selectSharedScenarios,executeAcceptance} from '../../scripts/gherkin.ts';
import {bindings,cleanup} from '../acceptance/steps.ts';
const path='workflows/docx/tracking-settings.feature';
const ids=['persistence','custody','no-op','refusal','rollback','plain-edit','author-refusal'].map(n=>'@id-docx-tracking-settings-'+n);
async function run(text:string,steps=bindings){const f=selectSharedScenarios(path,text,ids),count=(n:number)=>({implemented:n,planned:0,total:n});try{return await executeAcceptance({root:'.',features:[f],counts:{features:count(1),scenarios:count(7),cases:count(24),steps:count(f.scenarios.flatMap(s=>s.cases).reduce((n,c)=>n+c.steps.length,0))}},steps,'tracking-settings-outcomes');}finally{await cleanup();}}
test('all seven shared tracking settings contracts execute every one of 24 expanded cases',async()=>{
 const text=await Bun.file(join(fixturesRoot(),path)).text(),r=await run(text);expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(24);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);
});
test('incorrect tracking persistence and ordinary-edit expected values fail observable assertions',async()=>{
 const text=await Bun.file(join(fixturesRoot(),path)).text();for(const bad of [text.replace('the reopened tracking preference is <enabled>','the reopened tracking preference is false'),text.replace('the reopened text is Explicit plain edit','the reopened text is Wrong text')]){const r=await run(bad);expect(r.counts.cases.failed).toBeGreaterThan(0);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);}
});
