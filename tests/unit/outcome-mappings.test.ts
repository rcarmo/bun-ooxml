import {test,expect} from 'bun:test';
import {createHash} from 'node:crypto';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {reconcileOutcomeMappings,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
const source=`import {test,expect} from 'bun:test';test('saved outcome',()=>{expect(1).toBe(1)});test('unmapped',()=>{expect(2).toBe(2)});`;
const feature=`@planned
Feature: Saved outcome
 @id-pptx-mapping-test
 Scenario Outline: Saved <kind>
  Given input <kind>
  Then saved value matches
  Examples:
   | kind |
   | plain |
   | alias |
`;
function sample(){
 const records=inventoryTestSource('tests/unit/sample.test.ts',source);
 const inventory={cases:records};
 const features=[parseFeature('workflows/pptx/sample.feature',feature)];
 const sources={'tests/unit/sample.test.ts':source,'workflows/pptx/sample.feature':feature};
 const ledger:OutcomeMappingLedger={schemaVersion:1,consumer:'bun',executionCredit:false,scopePaths:['tests/unit/sample.test.ts'],sourceSha256:Object.fromEntries(Object.entries(sources).map(([p,s])=>[p,createHash('sha256').update(s).digest('hex')])),mappings:records.map(c=>({testId:c.id,status:'partial',scenarioIds:['@id-pptx-mapping-test'],caseKeys:[],assertions:[...c.assertions],outcomes:['A concrete saved-value assertion'],gaps:['Only this bounded assertion is mapped']}))};
 return {inventory,features,sources,ledger};
}
test('outcome mappings validate source pins and preserve the unmapped denominator without execution credit',()=>{
 const {inventory,features,sources,ledger}=sample();
 inventory.cases.push({...inventory.cases[0]!,id:'bun:elsewhere:unmapped',path:'tests/unit/other.test.ts'});
 const r=reconcileOutcomeMappings(inventory,ledger,features,sources) as any;
 expect(r.mappedDeclarations).toBe(2);expect(r.totalDeclarations).toBe(3);expect(r.unmappedTestIds).toEqual(['bun:elsewhere:unmapped']);expect(r.executionCredit).toBe(false);expect(r.mappings.every((m:any)=>m.executionCredit===false&&m.status==='partial')).toBe(true);
});
test('mapping gate refuses missing or duplicated test declarations',()=>{
 const {inventory,features,sources,ledger}=sample();ledger.mappings.pop();expect(()=>reconcileOutcomeMappings(inventory,ledger,features,sources)).toThrow('Missing');
 ledger.mappings.push(ledger.mappings[0]!);expect(()=>reconcileOutcomeMappings(inventory,ledger,features,sources)).toThrow('Duplicate');
});
test('mapping gate refuses stale source or nonexistent native assertions',()=>{
 const a=sample();a.sources['tests/unit/sample.test.ts']+=' ';expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('source');
 const b=sample();b.ledger.mappings[0]!.assertions=['expect(1).toBe(99)'];expect(()=>reconcileOutcomeMappings(b.inventory,b.ledger,b.features,b.sources)).toThrow('assertion');
});
test('mapping gate refuses unknown canonical scenarios, case keys and duplicate IDs',()=>{
 const a=sample();a.ledger.mappings[0]!.scenarioIds=['@id-absent'];expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('scenario');
 const b=sample();b.ledger.mappings[0]!.caseKeys=['@id-pptx-mapping-test|{"kind":"absent"}'];expect(()=>reconcileOutcomeMappings(b.inventory,b.ledger,b.features,b.sources)).toThrow('case');
 const c=sample();c.features.push(c.features[0]!);expect(()=>reconcileOutcomeMappings(c.inventory,c.ledger,c.features,c.sources)).toThrow('Duplicate');
});
test('mapping gate rejects missing source pins, invented scope paths and empty outcomes/gaps',()=>{
 const a=sample();delete a.ledger.sourceSha256['workflows/pptx/sample.feature'];expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('pin');
 const b=sample();b.ledger.scopePaths.push('tests/unit/missing.test.ts');expect(()=>reconcileOutcomeMappings(b.inventory,b.ledger,b.features,b.sources)).toThrow('scope');
 const c=sample();c.ledger.mappings[0]!.outcomes=[];expect(()=>reconcileOutcomeMappings(c.inventory,c.ledger,c.features,c.sources)).toThrow('outcome');
 const d=sample();d.ledger.mappings[0]!.gaps=[];expect(()=>reconcileOutcomeMappings(d.inventory,d.ledger,d.features,d.sources)).toThrow('gap');
});
test('mapping gate rejects fabricated execution credit and preserves dynamic review flags',()=>{
 const a=sample();(a.ledger as any).executionCredit=true;expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('credit');
 const b=sample();b.inventory.cases[0]!.reviewReasons.push('body-loop');
 const r=reconcileOutcomeMappings(b.inventory,b.ledger,b.features,b.sources) as any;
 expect(r.mappings[0].reviewReasons).toEqual(['body-loop']);expect(r.runtimeLeafCount).toBe(null);
});

test('committed mappings enumerate every slide-order test, assertions and known gaps',async()=>{
 const {outcomeMappingReport}=await import('../../scripts/outcome-mappings.ts');
 const report=await outcomeMappingReport();
 expect(report.mappedDeclarations).toBe(13);
 expect(report.unmappedTestIds.length).toBe(report.totalDeclarations-13);
 expect(report.mappings[0]!.caseKeys).toHaveLength(21);
 expect(report.mappings.every(m=>m.gaps.length>0&&m.outcomes.length>0&&m.assertions.length>0&&m.executionCredit===false)).toBe(true);
});

test('mapping gate refuses silently omitted native assertions',()=>{
 const a=sample();a.inventory.cases[0]!.assertions.push('expect(3).toBe(3)');
 expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('assertion');
});
test('scenario-only links remain explicitly distinct from case-key links',()=>{
 const a=sample();a.ledger.mappings[1]!.caseKeys=[a.features[0]!.scenarios[0]!.cases[0]!.identityKey];
 const r=reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources) as any;
 expect(r.mappings[0].linkGranularity).toBe('scenario-only');
 expect(r.mappings[1].linkGranularity).toBe('explicit-case-keys');
 expect(r.mappings.every((m:any)=>m.executionCredit===false&&m.status==='partial')).toBe(true);
});
