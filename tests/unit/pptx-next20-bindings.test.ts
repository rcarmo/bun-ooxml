import {test,expect} from 'bun:test';
import sources from '../../docs/behaviors/pptx-manipulation-next20-sources.json';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
import {parseFeature,executeAcceptance} from '../../scripts/gherkin.ts';
import {bindings,cleanup} from '../acceptance/pptx-next20.ts';
const path='features/pptx/manipulation-next20.feature';
export async function runPptxNext20(active=bindings){const feature=parseFeature(path,await Bun.file(path).text());const n=(v:number)=>({implemented:v,planned:0,total:v});try{return await executeAcceptance({root:'.',features:[feature],counts:{features:n(1),scenarios:n(20),cases:n(20),steps:n(100)}},active,'pptx-next20');}finally{await cleanup();}}
test('twenty PPTX manipulation cases run production saves, reopen and custody predicates',async()=>{const r=await runPptxNext20();expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(20);expect(r.counts.cases.total).toBe(20);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);expect(r.counts.steps.passed).toBe(100);});
test('twenty source capture identities and their immutable feature bytes remain separate from local execution credit',async()=>{
 expect(sources.executionCredit).toBe(false);expect(sources.records).toHaveLength(20);expect(new Set(sources.records.map(r=>r.sourceId)).size).toBe(20);
 const feature=parseFeature(path,await Bun.file(path).text());expect(feature.scenarios.map(s=>s.scenarioId)).toEqual(sources.records.map(r=>r.localId));
 for(const r of sources.records){const text=await Bun.file(join(fixturesRoot(),r.sourceFeature)).text();expect(new Bun.CryptoHasher('sha256').update(text).digest('hex')).toBe(r.sourceSha256);expect(text).toContain(r.sourceBlock);}
});
test('restored paragraph and custody predicates reject lost production edits and false caller copies',async()=>{
 for(const fault of ['lost-edit','caller-mutation']){const active=bindings.map(b=>b.pattern.test('the production manipulation is performed for patch-title')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{if(fault==='lost-edit'&&captures[0]==='patch-title')return;await b.run(c,...captures);if(fault==='caller-mutation')(c.state as {input:Uint8Array}).input[0]=0;}}:b);const r=await runPptxNext20(active);expect(r.counts.cases.failed).toBeGreaterThan(0);expect(r.counts.steps.ambiguous).toBe(0);expect(r.counts.steps.undefined).toBe(0);}
});
