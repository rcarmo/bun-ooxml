import {sharedScenarios} from '../helpers/shared-scenarios.ts';
import {expect,test} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {executeAcceptance,selectSharedScenarios,type StepBinding} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/steps.ts';
const ids=['@id-docx-go-new-empty-body','@id-docx-go-table-dimensions-getters','@id-docx-go-roundtrip-table-text'];
async function run(active:StepBinding[]=bindings,change:(s:string)=>string=s=>s){
 const count=(n:number)=>({implemented:n,planned:0,total:n});
 return executeAcceptance({root:'.',features:await sharedScenarios(ids,change),counts:{features:count(2),scenarios:count(3),cases:count(10),steps:count(33)}},active,'docx-model-unit');
}
function failed(r:Awaited<ReturnType<typeof run>>){expect(r.counts.cases.failed).toBeGreaterThan(0);expect(r.counts.steps.failed).toBeGreaterThan(0);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);}
test('ten existing DOCX document-model cases execute without activating incompatible table policies',async()=>{
 const result=await run();expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(10);
 for(const id of ['@id-docx-go-table-cell-access','@id-docx-go-table-row-counts','@id-docx-go-table-cell-text-getters','@id-docx-go-run-effects-getters'])expect(result.features.flatMap(f=>f.scenarios).some(s=>s.scenarioId===id)).toBe(false);
});
test('wrong table dimensions fail bound getter predicates',async()=>{
 failed(await run(bindings,s=>s.replace('RowCount equals <rows> and ColumnCount equals <cols> in memory','RowCount equals 99 and ColumnCount equals 99 in memory')));
});

test('model predicates reject a missing body and nonzero empty-document counts',async()=>{
 type State=import('../acceptance/docx-model.ts').ModelState;
 for(const mutate of [(s:State)=>{s.body=undefined;},(s:State)=>{s.paragraphCount=1;},(s:State)=>{s.tableCount=1;}]){
  const corrupt=bindings.map(b=>b.pattern.test('its body paragraphs and tables are enumerated')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);mutate(c.state as State);}}:b);
  failed(await run(corrupt));
 }
});

test('saved table checks inspect every cell position and reject changed delivered bytes',async()=>{
 type State=import('../acceptance/docx-model.ts').ModelState;
 for(let position=0;position<9;position++){
  const corrupt=bindings.map(b=>b.pattern.test('the document is saved and reopened')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);const s=c.state as State;s.reopened!.tables[0]!.cell(Math.floor(position/3),position%3).text='Wrong';}}:b);
  failed(await run(corrupt));
 }
 const corrupt=bindings.map(b=>b.pattern.test('the document is saved and reopened')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);const s=c.state as State;s.disk![0]=s.disk![0]!^1;}}:b);
 failed(await run(corrupt));
});

test('DOCX model path outputs are removed after successful and failed saves',async()=>{
 const {lstat}=await import('node:fs/promises');type State=import('../acceptance/docx-model.ts').ModelState;
 for(const refusal of [false,true]){
  let captured:State|undefined;
  const controlled=bindings.map(b=>b.pattern.test('the document is saved and reopened')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{
   captured=c.state as State;if(refusal)captured.document.save=async()=>{throw Error('injected path-save failure');};await b.run(c,...captures);
  }}:b);
  const r=await run(controlled);if(refusal)failed(r);else expect(r.failures).toEqual([]);
  expect(captured?.destination).toBeDefined();await expect(lstat(captured!.destination!)).rejects.toMatchObject({code:'ENOENT'});await expect(lstat(join(captured!.destination!,'..'))).rejects.toMatchObject({code:'ENOENT'});
 }
});

test('all eight dimension rows and the saved table exercise the requested grid sizes',async()=>{
 type State=import('../acceptance/docx-model.ts').ModelState;const dimensions:number[][]=[];
 const checking=bindings.map(b=>b.pattern.test('a table with 1 rows and 1 columns is added')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);const table=(c.state as State).document.tables[0]!;dimensions.push([table.rows,table.columns]);}}:b);
 expect((await run(checking)).failures).toEqual([]);expect(dimensions).toEqual([[1,1],[1,5],[5,1],[2,2],[3,3],[5,5],[10,3],[3,10]]);
});
