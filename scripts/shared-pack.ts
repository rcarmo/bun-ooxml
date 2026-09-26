import { join } from "node:path";
import { isDeepStrictEqual } from "node:util";
import { inventoryFeatures, parseFeature } from "./gherkin.ts";
import { verifyFixture, type FixtureManifest } from "./shared-fixtures.ts";

const ROOT=join(import.meta.dir,"..");
const PATH="docs/contracts/shared-v2";
const digest=(b:Uint8Array)=>new Bun.CryptoHasher("sha256").update(b).digest("hex");
const textHash=(s:string)=>digest(new TextEncoder().encode(s));
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

/** Verify only contract inventory, frozen inputs and native readback. This must
 * never emit a passed workflow outcome: all shared mutation bindings are planned.
 */
export async function verifySharedPack(root=ROOT):Promise<{scenarios:number;cases:number;fixtures:number;files:number}> {
  const metadata=await Bun.file(join(root,PATH,"manifest.json")).json();
  check(metadata.version===1&&metadata.workflowLayer===true&&metadata.upstreamApiParityCredit===false,"Workflow/library parity scope changed");
  check(safe(metadata.pack)&&safe(metadata.featurePath),"Unsafe shared pack path");
  const pack=join(root,metadata.pack);
  const manifestBytes=await Bun.file(join(pack,"pack-manifest.json")).bytes();
  check(digest(manifestBytes)===metadata.packManifestSha256,"Shared pack manifest drift");
  const manifest=JSON.parse(new TextDecoder().decode(manifestBytes));
  check(manifest.contractRevision===metadata.contractRevision&&manifest.scenarioCount===8&&manifest.expandedCaseCount===19&&manifest.lifecycle==="planned"&&manifest.bindingsImplemented===false,"Shared pack contract identity changed");
  const entries=Object.entries(manifest.files) as [string,string][];
  const known=new Set(["pack-manifest.json",...entries.map(([path])=>path)]);
  for(const [path,hash] of entries) {
    check(safe(path)&&/^[a-f0-9]{64}$/.test(hash),"Invalid shared file pin");
    check(digest(await Bun.file(join(pack,path)).bytes())===hash,`Shared pack artifact drift: ${path}`);
  }
  for await(const file of new Bun.Glob("**/*").scan({cwd:pack,onlyFiles:true}))check(known.has(file),`Unpinned shared pack artifact ${file}`);
  const featureText=await Bun.file(join(pack,"features/mutation-safety.feature")).text();
  check(metadata.execution?.lifecycle==='implemented'&&metadata.execution?.runner==='bun',"Missing execution-copy policy");
  const executionText=await Bun.file(join(root,metadata.featurePath)).text();
  check(executionText===featureText.replace(/^@planned/m,'@implemented @bun'),"Active shared feature differs beyond lifecycle tags");
  const executionFeature=parseFeature(metadata.featurePath,executionText);
  check(executionFeature.lifecycle==='implemented'&&executionFeature.runner==='bun',"Invalid native execution feature");
  const feature=parseFeature("features/mutation-safety.feature",featureText);
  check(feature.lifecycle==="planned"&&feature.scenarios.length===8,"Shared workflow coverage not planned");
  check(isDeepStrictEqual(feature.scenarios.map(s=>s.scenarioId),metadata.scenarios),"Shared scenario IDs changed");
  const fixtureManifest=await Bun.file(join(pack,"fixture-manifest.json")).json() as FixtureManifest;
  check(fixtureManifest.schemaVersion===1&&fixtureManifest.contractRevision===metadata.contractRevision&&fixtureManifest.fixtures.length===4,"Fixture inventory changed");
  check(safe(fixtureManifest.generator.path)&&fixtureManifest.generator.sha256===manifest.files[fixtureManifest.generator.path],"Fixture recipe drift");
  const source=await Bun.file(join(root,"references/manifest.json")).json();
  check(source.sources.some((s:any)=>s.id==="python-office-mcp-server"&&s.commit===metadata.sourceCommit),"Shared source commit differs from imported source");
  const fixtures=new Map<string,FixtureManifest["fixtures"][number]>();
  for(const fixture of fixtureManifest.fixtures) {
    check(safe(fixture.path)&&!fixtures.has(fixture.id),"Duplicate/unsafe fixture identity");fixtures.set(fixture.id,fixture);
    check(fixture.origin.revision===metadata.sourceCommit,"Fixture source revision drift");
    const original=source.files.find((f:any)=>f.source==="python-office-mcp-server"&&f.upstreamPath===fixture.origin.path);
    check(original&&original.sha256===fixture.origin.sha256,"Fixture origin not in pinned corpus");
    const fixtureBytes=await Bun.file(join(pack,fixture.path)).bytes();
    await verifyFixture(fixtureBytes,fixture);
    const allowed=new Set(fixture.allowedChangedPartsForSuccess);
    check(allowed.size===fixture.allowedChangedPartsForSuccess.length&&[...allowed].every(n=>Object.hasOwn(fixture.memberSha256,n)),"Invalid change allowance");
    const required=Object.fromEntries(Object.entries(fixture.memberSha256).filter(([n])=>!allowed.has(n)));
    check(isDeepStrictEqual(required,fixture.mustPreservePayloads),"Preservation allowance is not a complete partition");
  }
  const inventory=feature.scenarios.flatMap(s=>s.cases.map(c=>{
    const fixtureStep=c.steps.find(step=>/^fixture "/.test(step.text));
    const id=fixtureStep?.text.match(/^fixture "([^"]+)" verified against the fixture manifest$/)?.[1];
    const fixture=id?fixtures.get(id):undefined;check(fixture,"Case lacks a known fixture");
    for(const step of c.steps)if(step.argument?.dataTable)parseBatchTable(step.argument.dataTable);
    const examples=c.example?.values??{};
    return {scenarioId:s.scenarioId,stableCaseKey:stableCaseKey(s.scenarioId,examples),lifecycle:"planned",execution:"not-run",examples,featureSha256:textHash(featureText),fixture:{id,sha256:fixture.sha256,origin:fixture.origin},expandedSteps:c.steps.map(step=>({keyword:step.keyword,type:step.type,text:step.text,argument:step.argument??null,source:step.source}))};
  }));
  check(inventory.length===19 && new Set(inventory.map(c=>c.stableCaseKey)).size===19,"Missing/duplicate expanded cases");
  const expected={schemaVersion:1,contractRevision:metadata.contractRevision,validation:"parsed-compiled-and-typed-inputs-validated",scenarios:8,cases:19,bindingsImplemented:false,inventory};
  const expanded=await Bun.file(join(pack,"expanded-contracts.json")).json();
  // The parser includes optional undefined fields; JSON interchange omits them.
  // Compare the serialized envelope, without discarding any defined field/value.
  check(isDeepStrictEqual(JSON.parse(JSON.stringify(expected)),expanded),"Expanded steps/examples/source locations drifted");
  // Native execution copies receive new source hashes; shared stable keys and
  // step text/arguments remain exact. Passing status comes only from acceptance.
  check(isDeepStrictEqual(executionFeature.scenarios.map(s=>s.scenarioId),feature.scenarios.map(s=>s.scenarioId)),"Execution IDs drifted");
  const active=await inventoryFeatures(root);
  const historical=await Bun.file(join(root,"docs/contracts/office-mutation/manifest.json")).json();
  const activeIds=new Set(active.features.flatMap(f=>f.scenarios.map(s=>s.scenarioId)));
  for(const row of historical.scenarios)check(!activeIds.has(row.id)&&metadata.scenarios.includes(row.replacedBy),"Historical/new ID mapping missing or double-counted");
  return {scenarios:8,cases:19,fixtures:fixtures.size,files:entries.length};
}
if(import.meta.main)console.log("Shared pack input/inventory verification (no workflow execution):",await verifySharedPack());
