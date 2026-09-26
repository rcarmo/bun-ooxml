/** Source-pinned, bounded mapping integrity only; never awards execution credit. */
import {createHash} from 'node:crypto';
import {join} from 'node:path';
import {inventoryNativeTests,type TestCaseRecord} from './test-inventory.ts';
type NativeTestInventory={cases:TestCaseRecord[]};
import {parseFeature,type AcceptanceFeature} from './gherkin.ts';
import {fixturesRoot} from './fixture-inputs.ts';
export interface OutcomeMappingLedger {
 schemaVersion:1; consumer:'bun'; executionCredit:false;
 scopePaths:string[]; sourceSha256:Record<string,string>;
 mappings:Array<{testId:string; status:'partial'; scenarioIds:string[]; caseKeys:string[]; assertions:string[]; outcomes:string[]; gaps:string[]}>;
}
const sha=(text:string)=>createHash('sha256').update(text).digest('hex');
const nonempty=(values:string[])=>Array.isArray(values)&&values.length>0&&values.every(v=>typeof v==='string'&&v.trim().length>0)&&new Set(values).size===values.length;
export function reconcileOutcomeMappings(inventory:NativeTestInventory,ledger:OutcomeMappingLedger,features:AcceptanceFeature[],sources:Record<string,string>){
 if(ledger.schemaVersion!==1||ledger.consumer!=='bun'||ledger.executionCredit!==false)throw Error('Invalid mapping schema or execution credit');
 if(!nonempty(ledger.scopePaths))throw Error('Invalid mapping scope');
 const byId=new Map(inventory.cases.map(c=>[c.id,c]));
 if(byId.size!==inventory.cases.length)throw Error('Duplicate native declaration');
 for(const path of ledger.scopePaths)if(!inventory.cases.some(c=>c.path===path))throw Error('Unknown scope path: '+path);
 const scenarios=new Set<string>(),caseOwners=new Map<string,string>();
 for(const feature of features)for(const scenario of feature.scenarios){
  if(scenarios.has(scenario.scenarioId))throw Error('Duplicate canonical scenario: '+scenario.scenarioId);
  scenarios.add(scenario.scenarioId);
  for(const row of scenario.cases){if(caseOwners.has(row.identityKey))throw Error('Duplicate canonical case: '+row.identityKey);caseOwners.set(row.identityKey,scenario.scenarioId);}
 }
 // Callers provide the reviewed source set; omitting or adding a pin fails.
 // This bounded set is not a transitive runtime dependency closure.
 const required=new Set([...Object.keys(sources),...ledger.scopePaths,...features.map(f=>f.path)]);
 if([...required].some(p=>!Object.hasOwn(ledger.sourceSha256,p))||Object.keys(ledger.sourceSha256).some(p=>!required.has(p)))throw Error('Missing or unexpected source pin');
 for(const path of required)if(typeof sources[path]!=='string'||ledger.sourceSha256[path]!==sha(sources[path]!))throw Error('Stale source: '+path);
 const seen=new Set<string>();
 const mappings=ledger.mappings.map(mapping=>{
  if(seen.has(mapping.testId))throw Error('Duplicate mapped declaration: '+mapping.testId);seen.add(mapping.testId);
  const native=byId.get(mapping.testId);
  if(!native||!ledger.scopePaths.includes(native.path))throw Error('Unknown mapped declaration or scope: '+mapping.testId);
  if(mapping.status!=='partial')throw Error('Bounded mappings must retain partial status');
  if(!nonempty(mapping.scenarioIds)||mapping.scenarioIds.some(s=>!scenarios.has(s)))throw Error('Unknown or empty canonical scenario');
  if(!Array.isArray(mapping.caseKeys)||new Set(mapping.caseKeys).size!==mapping.caseKeys.length||mapping.caseKeys.some(k=>!mapping.scenarioIds.includes(caseOwners.get(k)!)))throw Error('Unknown or inconsistent canonical case');
  if(!nonempty(mapping.assertions)||mapping.assertions.length!==native.assertions.length||mapping.assertions.some(a=>!native.assertions.includes(a)))throw Error('Unknown, missing or empty native assertion');
  if(!nonempty(mapping.outcomes))throw Error('Missing concrete outcome');
  if(!nonempty(mapping.gaps))throw Error('Missing bounded gap');
  return {...mapping,linkGranularity:mapping.caseKeys.length?'explicit-case-keys':'scenario-only',reviewReasons:native.reviewReasons,deferredAssertions:native.deferredAssertions,unresolved:native.unresolved,executionCredit:false};
 });
 const missing=inventory.cases.filter(c=>ledger.scopePaths.includes(c.path)&&!seen.has(c.id));
 if(missing.length)throw Error('Missing scoped declaration: '+missing.map(c=>c.id).join(', '));
 return {schemaVersion:1,consumer:'bun',executionCredit:false,runtimeLeafCount:null,
  scope:'Source-pinned assertion mappings; semantic outcomes and gaps require human review; no test execution is inferred',
  totalDeclarations:inventory.cases.length,mappedDeclarations:mappings.length,sourceSha256:ledger.sourceSha256,
  unmappedTestIds:inventory.cases.filter(c=>!seen.has(c.id)).map(c=>c.id),mappings};
}
export async function outcomeMappingReport(){
 const root=process.cwd();
 const ledger=await Bun.file(join(root,'docs/behaviors/slide-order-mappings.json')).json() as OutcomeMappingLedger;
 const canonicalPath='workflows/pptx/slide-order.feature';
 const paths=['tests/unit/pptx-slide-order.test.ts','tests/acceptance/slide-order.ts','src/pptx/index.ts','src/pptx/slide-order.ts','scripts/gherkin.ts','scripts/test-inventory.ts'];
 const sources:Record<string,string>={};for(const path of paths)sources[path]=await Bun.file(join(root,path)).text();
 sources[canonicalPath]=await Bun.file(join(fixturesRoot(),canonicalPath)).text();
 return reconcileOutcomeMappings(await inventoryNativeTests(root),ledger,[parseFeature(canonicalPath,sources[canonicalPath])],sources);
}
if(import.meta.main){
 const output=JSON.stringify(await outcomeMappingReport(),null,2)+'\n',path='docs/behaviors/outcome-reconciliation.json';
 if(process.argv.includes('--check')){if(!await Bun.file(path).exists()||await Bun.file(path).text()!==output)throw Error('Outcome mapping report drift; review pins and regenerate');}
 else await Bun.write(path,output);
 const report=JSON.parse(output);console.log(`${report.mappedDeclarations}/${report.totalDeclarations} declarations have bounded mappings; ${report.unmappedTestIds.length} unmapped; no new execution credit`);
}
