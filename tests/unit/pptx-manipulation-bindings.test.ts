import {test,expect} from 'bun:test';
import {join} from 'node:path';
import pin from '../../docs/behaviors/pptx-manipulation-candidate.json';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {verifyReferences} from '../../scripts/references.ts';
import {selectSharedScenarios,executeAcceptance} from '../../scripts/gherkin.ts';
import {bindings,cleanup} from '../acceptance/pptx-manipulation.ts';
export async function runPptxManipulation(active=bindings){
 await verifyReferences(process.cwd());const root=fixturesRoot(),ledger=await Bun.file(join(root,'ledgers/pptx-manipulation.json')).json(),features=[];
 for(const path of new Set<string>(ledger.records.map((r:any)=>r.feature))){const ids=ledger.records.filter((r:any)=>r.feature===path).map((r:any)=>r.id),f=selectSharedScenarios('references/fixtures-ooxml/'+path,await Bun.file(join(root,path)).text(),ids);features.push({...f,scenarios:f.scenarios.filter(s=>ids.includes(s.scenarioId))});}
 const count=(v:number)=>({implemented:v,planned:0,total:v});try{return await executeAcceptance({root:process.cwd(),features,counts:{features:count(7),scenarios:count(20),cases:count(20),steps:count(179)}},active,'pptx-manipulation-shared');}finally{await cleanup();}
}
const enabled=!!process.env.OOXML_FIXTURES_ROOT&&!!process.env.OOXML_REFERENCE_PIN&&await Bun.file(process.env.OOXML_REFERENCE_PIN).json().then(p=>p.commit===pin.commit);
(enabled?test:test.skip)('sealed shared20PPTX cases execute179productionsteps with independent readback and custody',async()=>{const r=await runPptxManipulation();expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(20);expect(r.counts.steps.passed).toBe(179);for(const key of ['failed','undefined','ambiguous','skipped'] as const)expect(r.counts.steps[key]).toBe(0);});
(enabled?test:test.skip)('shared predicates reject lost editing and changed caller source',async()=>{for(const fault of ['lost-edit','caller-change']){const broken=bindings.map(b=>b.pattern.source.startsWith('^production presentation editing APIs')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{if(fault==='lost-edit')return;await b.run(c,...captures);(c.state as {input:Uint8Array}).input[0]=0;}}:b);const r=await runPptxManipulation(broken);expect(r.counts.cases.failed).toBeGreaterThan(0);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);}});
test('PPTX candidate retains complete feature seals and exact source/localidentity migrations without credit',()=>{expect(pin.schemaVersion).toBe(2);expect(pin.featureCount).toBe(64);expect(Object.keys(pin.featureSeals)).toHaveLength(64);expect(pin.selectedScenarioIds).toHaveLength(20);expect(pin.executionCredit).toBe(false);});
