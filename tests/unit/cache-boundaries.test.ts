import {beforeContract20Feature} from '../helpers/contract20-history.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {describe,expect,test} from 'bun:test';
import {join} from 'node:path';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {Workbook} from '../../src/xlsx/index.ts';
import {patchOffice} from '../../src/workflow/index.ts';
import {selectSharedScenarios,parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings,rangedFormulaFixture} from '../acceptance/cache-boundaries.ts';
describe('bounded formula cache custody',()=>{
 test('array and data-table inputs refuse without changing cached followers or earlier parts',async()=>{
  for(const kind of ['array','dataTable'] as const){const bytes=rangedFormulaFixture(kind);const w=await Workbook.open(bytes);
   expect(()=>w.worksheet('Model').setCellValue('A1',10)).toThrow('Cannot invalidate');expect(w.toBytes()).toEqual(bytes);
  }
 });
 test('workflow refuses and preserves an existing output when the workbook has array followers',async()=>{
  const root=await mkdtemp(join(tmpdir(),'bun-cache-scope-'));try{
   const source=join(root,'source.xlsx'),output=join(root,'output.xlsx'),bytes=rangedFormulaFixture('array');await Bun.write(source,bytes);await Bun.write(output,'old output');
   const r=await patchOffice({source,output,mode:'safe',changes:[{target:'Model!A1',value:10}]});
   expect(r.status).toBe('refused');expect(r.committedChanges).toBe(0);expect(r.error?.code).toBe('xlsx-cache-topology-unsupported');expect(Uint8Array.from(await Bun.file(source).bytes())).toEqual(Uint8Array.from(bytes));expect(await Bun.file(output).text()).toBe('old output');
  }finally{await rm(root,{recursive:true,force:true});}
 });
 test('executes every cache boundary Given/When/Then',async()=>{
  const {sharedScenarios}=await import('../helpers/shared-scenarios.ts');const path='workflows/xlsx/formula-cache.feature',f=selectSharedScenarios(path,await beforeContract20Feature(path,await Bun.file(join(fixturesRoot(),path)).text()),['@id-xlsx-array-input-refusal','@id-xlsx-cache-scope-opaque-parts']);f.scenarios=f.scenarios.filter(s=>['@id-xlsx-array-input-refusal','@id-xlsx-cache-scope-opaque-parts'].includes(s.scenarioId));if(!f)throw Error('Missing cache scenarios');
  const cases=f.scenarios.flatMap(s=>s.cases),steps=cases.reduce((n,c)=>n+c.steps.length,0);
  const count=(n:number)=>({implemented:n,planned:0,total:n});
  const inventory:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(cases.length),steps:count(steps)}};
  const r=await executeAcceptance(inventory,bindings,newAcceptanceRunId());expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(3);
 });
});
