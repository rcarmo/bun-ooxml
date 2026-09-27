import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {bindings,cleanup} from '../acceptance/steps.ts';
import {sharedScenarios} from '../helpers/shared-scenarios.ts';
import {executeAcceptance,inventoryFeatures} from '../../scripts/gherkin.ts';
import {outcomeMappingReport} from '../../scripts/outcome-mappings.ts';
const planned=['@id-python-word-anchor-section-discovery-hints','@id-docx-go-paragraph-style-getters','@id-docx-go-run-effects-getters','@id-docx-go-table-cell-access','@id-docx-go-table-merge-properties','@id-docx-go-track-author-toggle'];
async function scope(){const m=await Bun.file(join(fixturesRoot(),'ledgers/word-wording-migration.json')).json(),ids=m.files.flatMap((f:any)=>f.scenarios.map((s:any)=>s.id)) as string[],inv=await inventoryFeatures(process.cwd()),selected=inv.features.flatMap(f=>f.scenarios).filter(s=>ids.includes(s.scenarioId)&&s.lifecycle==='implemented').map(s=>s.scenarioId);return {m,ids,inv,selected};}
test('Word actor migration binds supported cases once without restoring old runtime names',async()=>{
 const {m,selected}=await scope();for(const file of m.files)for(const edit of file.stepRenames){expect(bindings.filter(b=>b.pattern.test(edit.from))).toHaveLength(0);if(selected.includes(edit.id))expect(bindings.filter(b=>b.pattern.test(edit.to))).toHaveLength(1);}
});
test('Word wording executes the same 28 supported IDs and 65 cases while six incompatible profiles stay planned',async()=>{
 const {ids,selected,inv}=await scope();expect(ids).toHaveLength(34);expect(selected).toHaveLength(28);expect(ids.filter(id=>!selected.includes(id)).sort()).toEqual([...planned].sort());
 expect(inv.features.flatMap(f=>f.scenarios).filter(s=>planned.includes(s.scenarioId)).reduce((n,s)=>n+s.cases.length,0)).toBe(10);
 const features=await sharedScenarios(selected),scenarios=features.flatMap(f=>f.scenarios),cases=scenarios.flatMap(s=>s.cases),count=(n:number)=>({implemented:n,planned:0,total:n});expect(cases).toHaveLength(65);
 try{const r=await executeAcceptance({root:'.',features,counts:{features:count(features.length),scenarios:count(28),cases:count(65),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}},bindings,'word-neutral');expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(65);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);}finally{await cleanup();}
});
test('Word actor changes preserve default execution identities and every prior mapping record',async()=>{
 const {inv}=await scope(),keys=inv.features.flatMap(f=>f.scenarios.filter(s=>s.lifecycle==='implemented').flatMap(s=>s.cases.map(c=>c.identityKey))).sort();expect(keys).toHaveLength(525);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(keys)).digest('hex')).toBe('cafa5b7815a43ebe12780b1d7b9742e0e806eeea2fdc4ba36f249d7927650d01');expect(inv.counts.cases.planned).toBe(35);
 const r=await outcomeMappingReport();expect(r.mappedDeclarations).toBe(472);expect(r.executionCredit).toBe(false);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(r.mappings)).digest('hex')).toBe('bc95c58ea9ff443697c4458ae733ed9a62eb53f436a948ee265cdd825a064742');
});
