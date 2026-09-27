import {test,expect} from 'bun:test';
import {join} from 'node:path';import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {selectSharedScenarios,executeAcceptance} from '../../scripts/gherkin.ts';
import {bindings,cleanup} from '../acceptance/steps.ts';
const path='workflows/docx/table-merging.feature',ids=['roundtrip','content-refusal','structure-refusal','coordinate-refusal','rollback','encoding','stale'].map(n=>'@id-docx-horizontal-merge-'+n);
async function run(source:string,steps=bindings){const feature=selectSharedScenarios(path,source,ids),count=(n:number)=>({implemented:n,planned:0,total:n});try{return await executeAcceptance({root:'.',features:[feature],counts:{features:count(1),scenarios:count(7),cases:count(26),steps:count(feature.scenarios.flatMap(s=>s.cases).reduce((n,c)=>n+c.steps.length,0))}},steps,'table-merge-outcomes');}finally{await cleanup();}}
test('seven horizontal merge contracts execute all26 expanded outcomes with no undefined or ambiguous steps',async()=>{
 const r=await run(await Bun.file(join(fixturesRoot(),path)).text());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(26);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);
});
test('wrong span and width outcomes fail concrete assertions',async()=>{
 const source=await Bun.file(join(fixturesRoot(),path)).text();for(const bad of [source.replace('gridSpan of <span>','gridSpan of 9'),source.replace('width is <width> twips','width is 1 twips')]){const r=await run(bad);expect(r.counts.cases.failed).toBe(3);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);}
});
test('corrupted saved text and unrelated payloads cannot pass physical merge custody checks',async()=>{
 const text=await Bun.file(join(fixturesRoot(),path)).text();for(const corrupt of [(s:any)=>s.reopened.paragraphs[0].setText('Wrong paragraph'),(s:any)=>s.reopened.package.setPart('customXml/opaque.bin',new Uint8Array([9]))]){const changed=bindings.map(b=>b.pattern.test('row 0 columns 0 through 2 are merged and saved to a new path')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);corrupt(c.merge);}}:b);const r=await run(text,changed);expect(r.counts.cases.failed).toBe(3);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);}
});
