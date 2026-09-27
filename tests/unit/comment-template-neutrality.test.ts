import {withoutTrackingToggle} from '../helpers/execution-baseline.ts';
import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {inventoryFeatures,parseFeature} from '../../scripts/gherkin.ts';
import {outcomeMappingReport} from '../../scripts/outcome-mappings.ts';
test('comment/template API labels do not activate unsupported consumer scenarios',async()=>{
 const m=await Bun.file(join(fixturesRoot(),'ledgers/comment-template-wording-migration.json')).json(),inv=await inventoryFeatures(process.cwd()),active=inv.features.flatMap(f=>f.scenarios),ids=m.files.flatMap((f:any)=>f.scenarios.map((r:any)=>r.id)) as string[];
 expect(ids).toHaveLength(24);const observed=active.filter(s=>ids.includes(s.scenarioId));expect(observed.filter(s=>s.lifecycle==='implemented').map(s=>s.scenarioId).sort()).toEqual(['@id-xml-escaping-values','@id-xml-typed-parse-error']);expect(observed.filter(s=>(s.lifecycle??'planned')==='planned')).toHaveLength(22);
 // Catalogue completeness makes templates visible but grants no execution.
 const templates=ids.filter(id=>id.includes('template'));expect(templates).toHaveLength(13);const loaded=inv.features.filter(f=>/\/template-(analysis|cache)\.feature$/.test(f.path));expect(loaded.flatMap(f=>f.scenarios)).toHaveLength(13);expect(loaded.every(f=>f.lifecycle==='planned'&&!f.runner)).toBe(true);
 for(const p of ['workflows/docx/template-analysis.feature','workflows/docx/template-cache.feature']){const f=parseFeature(p,await Bun.file(join(fixturesRoot(),p)).text());expect(f.lifecycle).toBe('planned');expect(f.scenarios.every(s=>(s.lifecycle??f.lifecycle)==='planned')).toBe(true);}
});
test('retained weak response and comment policies are not strengthened by neutral names',async()=>{
 const path='workflows/docx/template-analysis.feature',f=parseFeature(path,await Bun.file(join(fixturesRoot(),path)).text()),plain=f.scenarios.find(s=>s.scenarioId==='@id-python-word-template-analysis-plain-response')!;expect(plain.cases[0]!.steps.at(-1)!.text).toBe('its response is a dictionary');
 const p='workflows/docx/comments.feature',comments=parseFeature(p,await Bun.file(join(fixturesRoot(),p)).text());expect(comments.scenarios.find(s=>s.scenarioId==='@id-python-comments-filter-predicates')!.cases[0]!.steps.some(s=>s.text==='the open-filter, resolved-filter, and mine-filter results are each nonempty')).toBe(true);expect(comments.scenarios.find(s=>s.scenarioId==='@id-docx-comments-refusal')!.cases.some(c=>c.name.includes('missing-extension'))).toBe(true);
});
test('profile-only adoption preserves all default case identities and every bounded mapping record',async()=>{
 const inv=await inventoryFeatures(process.cwd()),keys=inv.features.flatMap(f=>f.scenarios.filter(s=>s.lifecycle==='implemented').flatMap(s=>s.cases.map(c=>c.identityKey))).sort();
 const originalKeys=withoutTrackingToggle(keys);expect(originalKeys).toHaveLength(525);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(originalKeys)).digest('hex')).toBe('cafa5b7815a43ebe12780b1d7b9742e0e806eeea2fdc4ba36f249d7927650d01');expect(inv.counts.cases.planned).toBe(47);
 const r=await outcomeMappingReport();expect(r.mappedDeclarations).toBe(494);expect(r.executionCredit).toBe(false);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(r.mappings.filter(m=>!['tracking-settings','heading-classification'].includes(m.ledger)))).digest('hex')).toBe('bc95c58ea9ff443697c4458ae733ed9a62eb53f436a948ee265cdd825a064742');
});
