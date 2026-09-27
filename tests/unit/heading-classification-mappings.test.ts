import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {inventoryNativeTests} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {outcomeMappingReport,reconcileOutcomeMappings,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
async function input(){const ledger=await Bun.file('docs/behaviors/heading-classification-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};for(const p of Object.keys(ledger.sourceSha256))sources[p]=await Bun.file(p.startsWith('workflows/')?join(fixturesRoot(),p):p).text();return {ledger,sources,inventory:await inventoryNativeTests(),features:[parseFeature('workflows/docx/paragraph-style.feature',sources['workflows/docx/paragraph-style.feature']!,{allowDuplicateCaseNames:true})]};}
test('eight direct-heading associations retain empty case keys and all486 prior mapping records',async()=>{
 const x=await input(),r=reconcileOutcomeMappings(x.inventory,x.ledger,x.features,x.sources);expect(r.mappedDeclarations).toBe(8);expect(r.mappings.every(m=>m.caseKeys.length===0&&m.status==='partial'&&m.gaps.length>0)).toBe(true);expect(r.executionCredit).toBe(false);
 const all=await outcomeMappingReport(),old=all.mappings.filter(m=>!['heading-classification','nullable-cell'].includes(m.ledger));expect(old).toHaveLength(486);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(old)).digest('hex')).toBe('b42eac023200eb822073f67568ea27047852d52f2e2be06d6adcc94812d34884');expect(all.mappedDeclarations).toBe(501);
});
test('heading mapping missing assertions stale runtime pins and unknown keys refuse',async()=>{
 const x=await input();for(const change of [(l:OutcomeMappingLedger)=>l.mappings.pop(),(l:OutcomeMappingLedger)=>l.mappings[0]!.assertions.pop(),(l:OutcomeMappingLedger)=>{l.sourceSha256['src/docx/index.ts']='0'.repeat(64);},(l:OutcomeMappingLedger)=>{l.mappings[0]!.caseKeys=['unknown'];}]){const l=structuredClone(x.ledger);change(l);expect(()=>reconcileOutcomeMappings(x.inventory,l,x.features,x.sources)).toThrow();}
});
