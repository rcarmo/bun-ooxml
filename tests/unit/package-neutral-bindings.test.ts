import {withoutTrackingToggle} from '../helpers/execution-baseline.ts';
import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {bindings,cleanup} from '../acceptance/steps.ts';
import {sharedScenarios} from '../helpers/shared-scenarios.ts';
import {executeAcceptance,inventoryFeatures} from '../../scripts/gherkin.ts';
import {outcomeMappingReport} from '../../scripts/outcome-mappings.ts';

test('neutral package actors resolve once and retired implementation names do not bind',async()=>{
 const migration=await Bun.file(join(fixturesRoot(),'ledgers/package-wording-migration.json')).json();
 for(const file of migration.files)for(const edit of file.stepRenames){expect(bindings.filter(b=>b.pattern.test(edit.to))).toHaveLength(1);expect(bindings.filter(b=>b.pattern.test(edit.from))).toHaveLength(0);}
});
test('all thirty-eight package and ZIP32 cases retain concrete assertions after actor migration',async()=>{
 const migration=await Bun.file(join(fixturesRoot(),'ledgers/package-wording-migration.json')).json(),ids=migration.files.flatMap((f:any)=>f.scenarios.map((s:any)=>s.id));
 const features=await sharedScenarios(ids),scenarios=features.flatMap(f=>f.scenarios),cases=scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});
 try{const r=await executeAcceptance({root:'.',features,counts:{features:count(2),scenarios:count(18),cases:count(38),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}},bindings,'package-neutral-wording');
 expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(38);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);
 }finally{await cleanup();}
});
test('package wording keeps default case identities all mapping records and planned gaps unchanged',async()=>{
 const inventory=await inventoryFeatures(process.cwd()),keys=inventory.features.flatMap(f=>f.scenarios.filter(s=>s.lifecycle==='implemented').flatMap(s=>s.cases.map(c=>c.identityKey))).sort();
 const originalKeys=withoutTrackingToggle(keys);
 expect(originalKeys).toHaveLength(525);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(originalKeys)).digest('hex')).toBe('cafa5b7815a43ebe12780b1d7b9742e0e806eeea2fdc4ba36f249d7927650d01');expect(inventory.counts.cases.planned).toBe(46);
 const report=await outcomeMappingReport();expect(report.mappedDeclarations).toBe(546);expect(report.executionCredit).toBe(false);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(report.mappings.filter(m=>!['tracking-settings','heading-classification','nullable-cell','table-merging','template-inventory'].includes(m.ledger)))).digest('hex')).toBe('bc95c58ea9ff443697c4458ae733ed9a62eb53f436a948ee265cdd825a064742');
});
