import {describe,expect,test} from "bun:test";
import {join} from "node:path";
import {cp,mkdir,mkdtemp,rm} from "node:fs/promises";
import {tmpdir} from "node:os";
import {parseBatchTable,stableCaseKey,verifySharedPack} from "../../scripts/shared-pack.ts";
import {parseFeature} from "../../scripts/gherkin.ts";
const root=join(import.meta.dir,"../..");
describe("shared v2 inventory and provenance",()=>{
  test("checks 8 scenarios, 19 cases and four native-readable pinned fixtures without executing workflows",async()=>{
    expect(await verifySharedPack(root)).toEqual({scenarios:8,cases:19,fixtures:4,files:17});
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

  test("rejects feature, fixture and compiled-inventory tampering while leaving original fixtures untouched",async()=>{
    const temp=await mkdtemp(join(tmpdir(),"bun-shared-pack-"));
    try {
      await mkdir(join(temp,"upstream"),{recursive:true});
      await Bun.write(join(temp,"references/manifest.json"),Bun.file(join(root,"references/manifest.json")));
      await cp(join(root,"features"),join(temp,"features"),{recursive:true});
      await cp(join(root,"docs/contracts/shared-v2"),join(temp,"docs/contracts/shared-v2"),{recursive:true});
      await cp(join(root,"docs/contracts/office-mutation"),join(temp,"docs/contracts/office-mutation"),{recursive:true});
      for(const [file,expectedError] of [
        ["features/workflow/mutation-safety.feature","Active shared feature differs"],
        ["docs/contracts/shared-v2/pack/fixtures/default-style.xlsx","Shared pack artifact drift"],
        ["docs/contracts/shared-v2/pack/expanded-contracts.json","Shared pack artifact drift"],
      ]) {
        const path=join(temp,file!);const original=await Bun.file(path).bytes();
        const modified=new Uint8Array(original.length+1);modified.set(original);modified[modified.length-1]=32;
        await Bun.write(path,modified);await expect(verifySharedPack(temp)).rejects.toThrow(expectedError!);await Bun.write(path,original);
      }
      expect(await verifySharedPack(temp)).toEqual({scenarios:8,cases:19,fixtures:4,files:17});
    } finally {await rm(temp,{recursive:true,force:true});}
  });
});
