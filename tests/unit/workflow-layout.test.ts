import {withoutTrackingToggle} from '../helpers/execution-baseline.ts';
import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {inventoryFeatures} from '../../scripts/gherkin.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {sharedScenarios} from '../helpers/shared-scenarios.ts';
import {outcomeMappingReport} from '../../scripts/outcome-mappings.ts';

test('format regrouping retains the exact 525 implemented shared case identities',async()=>{
 const inventory=await inventoryFeatures(process.cwd()),keys=inventory.features.flatMap(f=>f.scenarios.filter(s=>s.lifecycle==='implemented').flatMap(s=>s.cases.map(c=>c.identityKey))).sort();
 const originalKeys=withoutTrackingToggle(keys);
 expect(originalKeys).toHaveLength(525);expect(new Set(originalKeys).size).toBe(525);
 expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(originalKeys)).digest('hex')).toBe('cafa5b7815a43ebe12780b1d7b9742e0e806eeea2fdc4ba36f249d7927650d01');
 expect(inventory.counts.cases.planned).toBe(46);
 for(const f of inventory.features.filter(f=>f.path.startsWith('references/'))){expect(f.path).toMatch(/^references\/fixtures-ooxml\/workflows\/(docx|pptx|xlsx|package|xml|office)\/[a-z0-9-]+\.feature$/);}
});
test('multi-feature wrapper selection executes only requested identities and rejects missing ones',async()=>{
 const ids=['@id-office-preview-details','@id-office-docx-exact-match-counts'],features=await sharedScenarios(ids);
 expect(features).toHaveLength(2);expect(features.flatMap(f=>f.scenarios.map(s=>s.scenarioId))).toEqual(ids);
 await expect(sharedScenarios(['@id-not-in-catalogue'])).rejects.toThrow('one shared scenario owner');
 await expect(sharedScenarios([])).rejects.toThrow('nonempty unique');await expect(sharedScenarios([ids[0]!,ids[0]!])).rejects.toThrow('nonempty unique');
 const ledger=await Bun.file(join(fixturesRoot(),'ledgers/workflows.json')).json();
 for(const f of features)for(const s of f.scenarios)expect(ledger.workflows.find((w:any)=>w.id===s.scenarioId).feature).toBe(f.path);
});
test('mapping relocation preserves 470 records and explicitly accounts for two wrapper count changes',async()=>{
 const r=await outcomeMappingReport(),ids=['bun:tests/unit/pptx.test.ts:pptx slice / acceptance feature passes with the dedicated PPTX bindings','bun:tests/unit/docx-append-run.test.ts:two shared run-authoring cases execute through real save/reopen and reject corrupted saved formatting'];
 const unchanged=r.mappings.filter(m=>!['tracking-settings','heading-classification','nullable-cell','table-merging','template-inventory','comment-threads','revision-properties','revision-moves'].includes(m.ledger)&&!ids.includes(m.testId));expect(unchanged).toHaveLength(470);
 expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(unchanged)).digest('hex')).toBe('a87881cd02360bf6db2556d4cb39bc4a22bb59db90752d4fff83dd9f81960b41');
 expect(r.mappings.find(m=>m.testId===ids[0])!.assertions).toContain('expect(report.inventory.features.implemented).toBe(2)');
 expect(r.mappings.find(m=>m.testId===ids[1])!.assertions).toContain('expect(good.counts.cases.planned).toBe(0)');
 expect(r.mappedDeclarations).toBe(590);expect(r.executionCredit).toBe(false);expect(r.runtimeLeafCount).toBeNull();
});
