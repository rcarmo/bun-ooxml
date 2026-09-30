import {test,expect} from 'bun:test';
import {join,resolve} from 'node:path';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {verifyReferences} from '../../scripts/references.ts';
import {selectSharedScenarios,executeAcceptance,type StepBinding} from '../../scripts/gherkin.ts';
import {bindings,cleanup} from '../acceptance/steps.ts';
import candidate from '../../docs/behaviors/package-alignment-candidate.json';
const root=resolve(import.meta.dir,'../..');
const paths=['workflows/package/zip32.feature','workflows/package/zip-admission.feature','workflows/package/xml-member-admission.feature','workflows/package/preservation.feature','workflows/package/graph.feature','workflows/xml/comparison.feature','workflows/package/semantic-diff.feature','workflows/package/relationship-namespaces.feature'];
const enabled=!!process.env.OOXML_FIXTURES_ROOT&&!!process.env.OOXML_REFERENCE_PIN&&(await Bun.file(process.env.OOXML_REFERENCE_PIN).json()).commit===candidate.commit;
export async function runPackageAlignment(active:StepBinding[]=bindings,change=(text:string)=>text){
 await verifyReferences(root);const features=[];
 for(const path of paths){const text=change(await Bun.file(join(fixturesRoot(),path)).text()),ids=candidate.scenarioIds.filter(id=>text.includes(id+'\n'));const feature=selectSharedScenarios(path,text,ids);feature.scenarios=feature.scenarios.filter(s=>ids.includes(s.scenarioId));features.push(feature);}
 const n=(v:number)=>({implemented:v,planned:0,total:v});
 try{return await executeAcceptance({root,features,counts:{features:n(8),scenarios:n(20),cases:n(27),steps:n(127)}},active,'package-alignment');}finally{await cleanup();}
}
(enabled?test:test.skip)('twenty package IDs execute exactly27cases127steps including all71 corpus archives',async()=>{
 const r=await runPackageAlignment();expect(r.failures).toEqual([]);expect(r.counts.cases).toEqual({passed:27,failed:0,planned:0,total:27});expect(r.counts.steps.passed).toBe(127);expect(r.features.flatMap(f=>f.scenarios).map(s=>s.scenarioId).sort()).toEqual([...candidate.scenarioIds].sort());
});
(enabled?test:test.skip)('concrete package predicates reject wrong payloads, URI Type values and false custody without ambiguity',async()=>{
 for(const mutate of [
  (s:string)=>s.replace('payload 070809, effective type','payload 070800, effective type'),
  (s:string)=>s.replace('Target="b.xml" Type="urn:test/b" Id="b"','Target="b.xml" Type="urn:test/c" Id="b"'),
 ]){const r=await runPackageAlignment(bindings,mutate);expect(r.failures.length).toBeGreaterThan(0);}
 const active=bindings.map(b=>b.pattern.test('the production ZIP reader reads the archive with default limits')?{...b,run:async(c:Record<string,unknown>,...captures:string[])=>{await b.run(c,...captures);const s=c.state as {input:Uint8Array};s.input[0]=s.input[0]!^1;}}:b);
 const r=await runPackageAlignment(active);expect(r.counts.cases.failed).toBeGreaterThan(0);expect(r.counts.steps.ambiguous).toBe(0);expect(r.counts.steps.undefined).toBe(0);
});
(enabled?test:test.skip)('candidate wrong root, commit, manifest and feature identities refuse without altering shared files',async()=>{
 const tmp=await mkdtemp(join(tmpdir(),'package-candidate-pin-')),prior=process.env.OOXML_REFERENCE_PIN;
 try{const pin=join(tmp,'candidate.json');process.env.OOXML_REFERENCE_PIN=pin;
  for(const mutation of [{root:root},{root:undefined},{featureSeals:undefined},{featureCount:61},{commit:'0'.repeat(40)},{manifestSha256:'0'.repeat(64)},{featureSeals:{...candidate.featureSeals,'workflows/package/zip32.feature':'0'.repeat(64)}}]){
   await Bun.write(pin,JSON.stringify({...candidate,...mutation}));await expect(verifyReferences(root)).rejects.toThrow();
  }
 }finally{process.env.OOXML_REFERENCE_PIN=prior;await rm(tmp,{recursive:true,force:true});}
 expect(await verifyReferences(root)).toBe(346);
});
test('candidate contract retains complete feature seals and exact reviewed Office case-identity migrations without credit',()=>{
 expect(candidate.schemaVersion).toBe(2);expect(candidate.featureCount).toBe(62);expect(candidate.executionCredit).toBe(false);expect(candidate.cases).toBe(27);expect(candidate.steps).toBe(127);expect(candidate.scenarioIds).toHaveLength(20);expect(Object.keys(candidate.featureSeals)).toHaveLength(62);expect(candidate.caseIdentityMigration).toHaveLength(4);expect(candidate.caseIdentityMigration.every(r=>r.after!==r.before)).toBe(true);
});
