import {describe,expect,test} from "bun:test";
import {join} from "node:path";
import {cp,mkdir,mkdtemp,rm} from "node:fs/promises";
import {tmpdir} from "node:os";
import {parseBatchTable,stableCaseKey,verifySharedContracts} from "../../scripts/shared-contracts.ts";
import {inventoryFeatures,parseFeature} from "../../scripts/gherkin.ts";
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
const root=join(import.meta.dir,"../..");
describe("shared mutation inventory and custody",()=>{
  test("checks 8 scenarios, 19 cases and four native-readable pinned fixtures without executing workflows",async()=>{
    expect(await verifySharedContracts(root)).toEqual({scenarios:8,cases:19,fixtures:4,files:2});
    const ledger=await Bun.file(join(fixturesRoot(),'ledgers/workflows.json')).json();
    const canonical=ledger.features.filter((path:string)=>path.startsWith('workflows/native/'));
    expect(canonical).toHaveLength(8);
    const inventory=await inventoryFeatures(root);
    for(const path of canonical){
      const shared=inventory.features.filter(f=>f.path==='references/fixtures-ooxml/'+path);
      expect(shared).toHaveLength(1);
      expect(shared[0]!.lifecycle).toBe('implemented');
      const ids=parseFeature(path,await Bun.file(join(fixturesRoot(),path)).text()).scenarios.map(s=>s.scenarioId);
      for(const id of ids)expect(inventory.features.flatMap(f=>f.scenarios).filter(s=>s.scenarioId===id)).toHaveLength(1);
    }
  });
  test("stable identity excludes path, UUID and example insertion order",()=>{
    expect(stableCaseKey("@id-example",{z:"2",a:"1"})).toBe('@id-example:{"a":"1","z":"2"}');
    expect(stableCaseKey("@id-example")).toBe('@id-example:{}');
    expect(stableCaseKey("@id-example",{a:"1",z:"2"})).toBe(stableCaseKey("@id-example",{z:"2",a:"1"}));
  });
  test("typed mutation inputs preserve strings, numeric values and significant whitespace",()=>{
    expect(parseBatchTable([["target","value_json"],["Input!A1","10"],["A2",'"10"'],["A3",'"first\\nsecond"'],["A4","null"]])).toEqual([
      {target:"Input!A1",value:10},{target:"A2",value:"10"},{target:"A3",value:"first\nsecond"},{target:"A4",value:null}
    ]);
    for(const table of [[],[["target","value_json"]],[["target","value"],["A1","10"]],[["target","value_json"],["A1","not-json"]],[["target","value_json"],["A1","{}"]],[["target","value_json"],["A1","1e999"]]])expect(()=>parseBatchTable(table)).toThrow();
  });
  test("Gherkin escaping must preserve JSON newline escapes before typed decoding",()=>{
    const compile=(value:string)=>parseFeature('features/planned/example.feature',`@planned
Feature: JSON escaping
  @id-json-escaping
  Scenario: Typed newline
    When a batch is prepared:
      | target | value_json |
      | A1 | ${value} |
    Then the typed value contains a newline
`).scenarios[0]!.cases[0]!.steps[0]!.argument!.dataTable!;
    expect(()=>parseBatchTable(compile('"first\\nsecond"'))).toThrow();
    expect(parseBatchTable(compile('"first\\\\nsecond"'))).toEqual([{target:'A1',value:'first\nsecond'}]);
  });

  test("rejects feature, fixture and policy tampering while leaving original fixtures untouched",async()=>{
    const temp=await mkdtemp(join(tmpdir(),"bun-shared-contracts-"));
    try {
      await cp(join(root,"features"),join(temp,"features"),{recursive:true});
      await cp(fixturesRoot(),join(temp,"references/fixtures-ooxml"),{recursive:true,dereference:true,filter:(path)=>!path.includes('/node_modules')&&!path.endsWith('/.git')});
      const manifest=await Bun.file(join(temp,'references/fixtures-ooxml/manifest.json')).json();
      const contract=await Bun.file(join(temp,'references/fixtures-ooxml/contracts/mutation-safety.json')).json();
      const fixture=manifest.files.find((f:any)=>f.id===contract.fixtures[0].assetId);
      for(const [file,expectedError] of [
        ["references/fixtures-ooxml/workflows/mutation-safety.feature","Shared contract artifact drift"],
        ['references/fixtures-ooxml/'+fixture.path,'sha256 drift'],
        ["references/fixtures-ooxml/contracts/mutation-safety.json","Shared contract artifact drift"],
      ]) {
        const path=join(temp,file!);const original=await Bun.file(path).bytes();
        const modified=new Uint8Array(original.length+1);modified.set(original);modified[modified.length-1]=32;
        await Bun.write(path,modified);await expect(verifySharedContracts(temp)).rejects.toThrow(expectedError!);await Bun.write(path,original);
      }
      expect(await verifySharedContracts(temp)).toEqual({scenarios:8,cases:19,fixtures:4,files:2});
    } finally {await rm(temp,{recursive:true,force:true});}
  });
});
