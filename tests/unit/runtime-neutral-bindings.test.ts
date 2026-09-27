import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {bindings} from '../acceptance/steps.ts';
import {sharedScenarios} from '../helpers/shared-scenarios.ts';
import {executeAcceptance,inventoryFeatures} from '../../scripts/gherkin.ts';
import {outcomeMappingReport} from '../../scripts/outcome-mappings.ts';

test('neutral action wording has exactly one binding and retired actor wording has none',async()=>{
 const migration=await Bun.file(join(fixturesRoot(),'ledgers/runtime-wording-migration.json')).json();
 for(const file of migration.files)for(const edit of file.stepRenames){
  expect(bindings.filter(b=>b.pattern.test(edit.to))).toHaveLength(1);
  expect(bindings.filter(b=>b.pattern.test(edit.from))).toHaveLength(0);
 }
});
test('all eighteen portable-profile cases execute the retained inputs and outcomes',async()=>{
 const migration=await Bun.file(join(fixturesRoot(),'ledgers/runtime-wording-migration.json')).json(),ids=migration.files.flatMap((f:any)=>f.scenarios.map((s:any)=>s.id));
 const features=await sharedScenarios(ids),scenarios=features.flatMap(f=>f.scenarios),cases=scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 const r=await executeAcceptance({root:'.',features,counts:{features:count(4),scenarios:count(16),cases:count(18),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}},bindings,'neutral-wording');
 expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(18);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);
});
test('wording and profile labels do not change default acceptance identities or credit',async()=>{
 const inventory=await inventoryFeatures(process.cwd()),keys=inventory.features.flatMap(f=>f.scenarios.filter(s=>s.lifecycle==='implemented').flatMap(s=>s.cases.map(c=>c.identityKey))).sort();
 expect(keys).toHaveLength(525);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(keys)).digest('hex')).toBe('cafa5b7815a43ebe12780b1d7b9742e0e806eeea2fdc4ba36f249d7927650d01');expect(inventory.counts.cases.planned).toBe(35);
 const report=await outcomeMappingReport();expect(report.mappedDeclarations).toBe(472);expect(report.executionCredit).toBe(false);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(report.mappings)).digest('hex')).toBe('bc95c58ea9ff443697c4458ae733ed9a62eb53f436a948ee265cdd825a064742');
});
