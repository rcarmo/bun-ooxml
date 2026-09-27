import { join } from "node:path";
import { isDeepStrictEqual } from "node:util";
import { inventoryFeatures, parseFeature } from "./gherkin.ts";
import { verifyFixture, loadMutationFixtures } from "./shared-fixtures.ts";

const ROOT=join(import.meta.dir,"..");
const digest=(b:Uint8Array)=>new Bun.CryptoHasher("sha256").update(b).digest("hex");
function check(condition:unknown,message:string):asserts condition {if(!condition)throw new Error(message);}
function safe(path:unknown):path is string {return typeof path==="string"&&!!path&&!/[\\:\u0000-\u001f]/.test(path)&&path.split("/").every(p=>!!p&&p!=="."&&p!=="..");}

/** Cross-runtime case identity excludes feature file paths and compiler UUIDs.
 * Values stay strings as defined by Gherkin Examples. Mutation data-table values
 * are parsed separately; coercing example values would change contract identity.
 */
export function stableCaseKey(scenarioId:string,examples:Record<string,string>={}):string {
  check(/^@id-[a-z0-9-]+$/.test(scenarioId),"Invalid scenario ID");
  check(Object.values(examples).every(v=>typeof v==="string"),"Examples must be strings");
  return scenarioId+":"+JSON.stringify(Object.fromEntries(Object.entries(examples).sort(([a],[b])=>a.localeCompare(b))));
}

/** Decode typed workflow inputs. This helper does not apply edits or fabricate
 * receipts. A malformed table refuses before any caller can start mutation.
 */
export function parseBatchTable(table:string[][]):{target:string;value:string|number|boolean|null}[] {
  check(isDeepStrictEqual(table[0],["target","value_json"]),"Expected target/value_json table");
  check(table.length>1,"Empty mutation batch");
  return table.slice(1).map(row=>{
    check(row.length===2 && !!row[0],"Malformed batch row");
    const value:unknown=JSON.parse(row[1]!);
    check(value===null||typeof value==="string"||typeof value==="boolean"||(typeof value==="number"&&Number.isFinite(value)),"Unsupported batch value");
    return {target:row[0]!,value:value as string|number|boolean|null};
  });
}

/** Validate the canonical workflow and fixture policy; execution belongs to acceptance. */
export async function verifySharedContracts(root=ROOT):Promise<{scenarios:number;cases:number;fixtures:number;files:number}> {
 const referenceRoot=root===ROOT ? (process.env.OOXML_FIXTURES_ROOT??join(root,'references/fixtures-ooxml')) : join(root,'references/fixtures-ooxml');
 const manifest=await Bun.file(join(referenceRoot,'manifest.json')).json();
 const contractPath='contracts/mutation-safety.json';
 const verifySeal=async(path:string)=>{
  const entry=manifest.files.find((f:any)=>f.path===path);check(entry,'Missing shared contract manifest entry');
  check(digest(await Bun.file(join(referenceRoot,path)).bytes())===entry.sha256,'Shared contract artifact drift: '+path);
 };
 await verifySeal(contractPath);
 const contract=await Bun.file(join(referenceRoot,contractPath)).json();
 check(contract.schemaVersion===2&&!('feature' in contract)&&Array.isArray(contract.features)&&contract.features.length>0&&new Set(contract.features).size===contract.features.length,'Invalid mutation feature list');
 check(contract.scenarioIds.length===8&&new Set(contract.scenarioIds).size===8&&contract.expandedCaseCount===19,'Mutation inventory changed');
 const features=[];
 for(const path of contract.features){
  check(safe(path)&&/^workflows\/(docx|pptx|xlsx)\/[a-z-]+\.feature$/.test(path),'Unsafe mutation feature path');await verifySeal(path);
  const feature=parseFeature(path,await Bun.file(join(referenceRoot,path)).text(),{allowDuplicateCaseNames:true});
  check(feature.lifecycle==='planned','Canonical scenario lifecycle changed');
  const scenarios=feature.scenarios.filter(s=>contract.scenarioIds.includes(s.scenarioId));check(scenarios.length,'Unused mutation feature');features.push({...feature,scenarios});
 }
 const scenarios=features.flatMap(f=>f.scenarios),featureCases=scenarios.flatMap(s=>s.cases);
 check(isDeepStrictEqual(scenarios.map(s=>s.scenarioId).sort(),[...contract.scenarioIds].sort()),'Canonical scenario IDs/lifecycle changed');
 check(featureCases.length===contract.expandedCaseCount,'Expanded mutation case count changed');
 const fixtures=(await loadMutationFixtures(referenceRoot)).fixtures;
 check(fixtures.length===4,'Mutation fixture policy count changed');
 for(const fixture of fixtures){
  check(safe(fixture.path),'Unsafe fixture path');await verifyFixture(await Bun.file(join(referenceRoot,fixture.path)).bytes(),fixture);
 }
 const keys=new Set<string>();
 for(const scenario of scenarios)for(const c of scenario.cases){
  const key=stableCaseKey(scenario.scenarioId,c.example?.values??{});check(!keys.has(key),'Duplicate shared case key');keys.add(key);
  const id=c.steps.find(s=>/^fixture "/.test(s.text))?.text.match(/^fixture "([^"]+)" verified against the fixture manifest$/)?.[1];
  check(fixtures.some(f=>f.id===id),'Unknown workflow fixture');
  for(const step of c.steps)if(step.argument?.dataTable)parseBatchTable(step.argument.dataTable);
 }
 const config=await Bun.file(join(root,'features/shared.json')).json();
 for(const feature of features)check(config.features.some((f:any)=>f.path==='references/fixtures-ooxml/'+feature.path&&f.lifecycle==='implemented'&&f.runner==='bun'&&feature.scenarios.every(s=>f.scenarioIds.includes(s.scenarioId))),'Missing workflow execution mapping');
 const active=await inventoryFeatures(root),ids=active.features.flatMap(f=>f.scenarios.map(s=>s.scenarioId));
 for(const id of contract.scenarioIds)check(ids.filter(x=>x===id).length===1,'Workflow must occur once in execution inventory');
 return {scenarios:8,cases:keys.size,fixtures:fixtures.length,files:1+features.length};
}
if(import.meta.main)console.log('Shared workflow input/inventory verification (no execution):',await verifySharedContracts());
