import {withoutTrackingToggle} from '../helpers/execution-baseline.ts';
import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {bindings} from '../acceptance/steps.ts';
import {sharedScenarios} from '../helpers/shared-scenarios.ts';
import {executeAcceptance,inventoryFeatures} from '../../scripts/gherkin.ts';
import {outcomeMappingReport} from '../../scripts/outcome-mappings.ts';
test('neutral XML and formula actors have one handler with no retired runtime alias',async()=>{
 const m=await Bun.file(join(fixturesRoot(),'ledgers/xml-formula-wording-migration.json')).json();for(const f of m.files)for(const e of f.stepRenames){expect(bindings.filter(b=>b.pattern.test(e.to))).toHaveLength(1);expect(bindings.filter(b=>b.pattern.test(e.from))).toHaveLength(0);}
});
test('all sixty retained XML and formula cases execute exact lexical and static-reference predicates',async()=>{
 const m=await Bun.file(join(fixturesRoot(),'ledgers/xml-formula-wording-migration.json')).json(),ids=m.files.flatMap((f:any)=>f.scenarios.map((r:any)=>r.id)),features=await sharedScenarios(ids),cases=features.flatMap(f=>f.scenarios.flatMap(s=>s.cases)),count=(n:number)=>({implemented:n,planned:0,total:n});
 const r=await executeAcceptance({root:'.',features,counts:{features:count(2),scenarios:count(19),cases:count(60),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}},bindings,'xml-formula-neutral');expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(60);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);
});
test('XML formula actor migration preserves all default case identities and bounded mapping records',async()=>{
 const inv=await inventoryFeatures(process.cwd()),keys=inv.features.flatMap(f=>f.scenarios.filter(s=>s.lifecycle==='implemented').flatMap(s=>s.cases.map(c=>c.identityKey))).sort();
 const originalKeys=withoutTrackingToggle(keys);expect(originalKeys).toHaveLength(525);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(originalKeys)).digest('hex')).toBe('cafa5b7815a43ebe12780b1d7b9742e0e806eeea2fdc4ba36f249d7927650d01');expect(inv.counts.cases.planned).toBe(46);
 const r=await outcomeMappingReport();expect(r.mappedDeclarations).toBe(516);expect(r.executionCredit).toBe(false);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(r.mappings.filter(m=>!['tracking-settings','heading-classification','nullable-cell','table-merging'].includes(m.ledger)))).digest('hex')).toBe('bc95c58ea9ff443697c4458ae733ed9a62eb53f436a948ee265cdd825a064742');
});
