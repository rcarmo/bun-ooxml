/** Source-pinned, bounded mapping integrity only; never awards execution credit. */
import {createHash} from 'node:crypto';
import {join} from 'node:path';
import {inventoryNativeTests,type TestCaseRecord} from './test-inventory.ts';
type NativeTestInventory={cases:TestCaseRecord[];unresolved:readonly unknown[]};
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
 if(!Array.isArray(inventory.unresolved)||inventory.unresolved.length||inventory.cases.some(c=>!Array.isArray(c.unresolved)||c.unresolved.length))throw Error('Unresolved native registrations prevent outcome reconciliation');
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
export type OutcomeMappingSet={name:string;expectedScopePaths:string[];ledger:OutcomeMappingLedger;features:AcceptanceFeature[];sources:Record<string,string>};
export function reconcileOutcomeMappingSets(inventory:NativeTestInventory,sets:OutcomeMappingSet[]) {
 if(!sets.length)throw Error('At least one mapping ledger is required');
 const names=new Set<string>(),scopes=new Set<string>(),scenarios=new Set<string>();
 const sourceSha256:Record<string,string>={};
 const reports=sets.map(set=>{
  if(!set.name.trim()||names.has(set.name))throw Error('Duplicate or empty ledger name');names.add(set.name);
  if(!nonempty(set.expectedScopePaths)||set.ledger.scopePaths.length!==set.expectedScopePaths.length||set.expectedScopePaths.some(p=>!set.ledger.scopePaths.includes(p)))throw Error('Unexpected ledger scope');
  for(const path of set.ledger.scopePaths){if(scopes.has(path))throw Error('Duplicate ledger scope: '+path);scopes.add(path);}
  // Each ledger validates only against its own canonical features and source pins.
  const report=reconcileOutcomeMappings(inventory,set.ledger,set.features,set.sources);
  for(const feature of set.features)for(const scenario of feature.scenarios){
   if(scenarios.has(scenario.scenarioId))throw Error('Duplicate cross-ledger scenario: '+scenario.scenarioId);scenarios.add(scenario.scenarioId);
  }
  for(const [path,hash]of Object.entries(report.sourceSha256)){
   if(Object.hasOwn(sourceSha256,path)&&sourceSha256[path]!==hash)throw Error('Conflicting shared source: '+path);sourceSha256[path]=hash;
  }
  return {name:set.name,report};
 });
 const mappings=reports.flatMap(({name,report})=>report.mappings.map(mapping=>({...mapping,ledger:name}))),seen=new Set(mappings.map(m=>m.testId));
 if(seen.size!==mappings.length)throw Error('Duplicate cross-ledger declaration');
 return {schemaVersion:2,consumer:'bun',executionCredit:false,runtimeLeafCount:null,
  scope:'Independent source-pinned partial ledgers; integrity only, no test execution or semantic completeness inferred',
  totalDeclarations:inventory.cases.length,mappedDeclarations:mappings.length,sourceSha256,
  ledgers:reports.map(({name,report})=>({name,scopePaths:sets.find(s=>s.name===name)!.ledger.scopePaths,mappedDeclarations:report.mappedDeclarations,sourceSha256:report.sourceSha256})),
  unmappedTestIds:inventory.cases.filter(c=>!seen.has(c.id)).map(c=>c.id),mappings};
}
export async function outcomeMappingReport(){
 const root=process.cwd();
 // Fixed registrations prevent removing a ledger or shrinking its scope in JSON.
 const registrations=[
  {name:'slide-order',tests:['tests/unit/pptx-slide-order.test.ts'],canonical:'workflows/pptx/slide-order.feature',sources:['tests/acceptance/slide-order.ts','src/pptx/index.ts','src/pptx/slide-order.ts']},
  {name:'effective-formatting',tests:['tests/unit/docx-effective-formatting.test.ts'],canonical:'workflows/docx/effective-formatting.feature',sources:['tests/acceptance/effective-formatting.ts','src/docx/index.ts','src/docx/effective-formatting.ts']},
  {name:'xml-values',tests:['tests/unit/xml.test.ts'],canonical:'workflows/xml/parsing.feature',sources:['tests/acceptance/core.ts','tests/acceptance/xml-values.ts','src/xml/index.ts','src/errors.ts']},
  {name:'pptx-core',tests:['tests/unit/pptx.test.ts'],canonical:'workflows/native/pptx-text.feature',sources:['tests/acceptance/pptx.ts','tests/acceptance/pptx-custody.ts','src/pptx/index.ts','src/opc/package.ts','src/opc/zip.ts','scripts/acceptance.ts','scripts/fixture-inputs.ts']},
  {name:'docx-model',tests:['tests/unit/docx-append-run.test.ts','tests/unit/docx-cell-properties.test.ts','tests/unit/docx-row-header.test.ts','tests/unit/docx-paragraph-text.test.ts','tests/unit/docx-insert-paragraph.test.ts','tests/unit/docx-table-rows.test.ts','tests/unit/docx-row-texts.test.ts'],canonical:'workflows/docx/document-model.feature',sources:['tests/acceptance/append-run.ts','tests/acceptance/cell-properties.ts','tests/acceptance/row-header.ts','tests/acceptance/docx-model.ts','tests/acceptance/steps.ts','tests/acceptance/paragraph-text.ts','tests/acceptance/paragraph-properties.ts','tests/acceptance/body-insertion.ts','tests/acceptance/table-rows.ts','tests/acceptance/row-texts.ts','src/docx/index.ts','src/docx/append-run.ts','src/docx/cell-properties.ts','src/docx/row-header.ts','src/docx/paragraph-text.ts','src/docx/body-insertion.ts','src/docx/table-rows.ts','src/docx/run-formatting.ts','src/docx/paragraph-style.ts','src/opc/package.ts','src/xml/index.ts']},
  {name:'formula-references',tests:['tests/unit/xlsx-range.test.ts','tests/unit/xlsx-formula-analysis.test.ts','tests/unit/xlsx-formula-remap.test.ts'],canonical:'workflows/xlsx/formula-references.feature',sources:['src/xlsx/range.ts','src/xlsx/formula.ts','src/xlsx/formula-remap.ts','src/errors.ts','tests/acceptance/xlsx-range.ts','tests/acceptance/formula-analysis.ts','tests/acceptance/formula-remap.ts']},
 ];
 const sets:OutcomeMappingSet[]=[];
 for(const registration of registrations){
  const ledger=await Bun.file(join(root,`docs/behaviors/${registration.name}-mappings.json`)).json() as OutcomeMappingLedger;
  const sources:Record<string,string>={};for(const path of [...registration.tests,...registration.sources,'scripts/gherkin.ts','scripts/test-inventory.ts'])sources[path]=await Bun.file(join(root,path)).text();
  sources[registration.canonical]=await Bun.file(join(fixturesRoot(),registration.canonical)).text();
  sets.push({name:registration.name,expectedScopePaths:registration.tests,ledger,features:[parseFeature(registration.canonical,sources[registration.canonical]!,{allowDuplicateCaseNames:true})],sources});
 }
 return reconcileOutcomeMappingSets(await inventoryNativeTests(root),sets);
}
if(import.meta.main){
 const output=JSON.stringify(await outcomeMappingReport(),null,2)+'\n',path='docs/behaviors/outcome-reconciliation.json';
 if(process.argv.includes('--check')){if(!await Bun.file(path).exists()||await Bun.file(path).text()!==output)throw Error('Outcome mapping report drift; review pins and regenerate');}
 else await Bun.write(path,output);
 const report=JSON.parse(output);console.log(`${report.mappedDeclarations}/${report.totalDeclarations} declarations have bounded mappings; ${report.unmappedTestIds.length} unmapped; no new execution credit`);
}
