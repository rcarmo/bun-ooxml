import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {inventoryNativeTests} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {outcomeMappingReport,reconcileOutcomeMappings,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
async function input(){const ledger=await Bun.file('docs/behaviors/tracking-settings-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};for(const p of Object.keys(ledger.sourceSha256))sources[p]=await Bun.file(p.startsWith('workflows/')?join(fixturesRoot(),p):p).text();return {ledger,sources,inventory:await inventoryNativeTests(),features:Object.entries(sources).filter(([p])=>p.startsWith('workflows/')).map(([p,s])=>parseFeature(p,s,{allowDuplicateCaseNames:true}))};}
test('tracking mappings cover fourteen declarations with bounded case keys and preserve all472 older records',async()=>{
 const x=await input(),r=reconcileOutcomeMappings(x.inventory,x.ledger,x.features,x.sources);expect(r.mappedDeclarations).toBe(14);expect(r.mappings.every(m=>m.status==='partial'&&m.gaps.length>0)).toBe(true);expect(r.mappings.filter(m=>m.caseKeys.length).map(m=>m.caseKeys.length)).toEqual([1,24]);
 const all=await outcomeMappingReport(),old=all.mappings.filter(m=>!['tracking-settings','heading-classification'].includes(m.ledger));expect(old).toHaveLength(472);expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(old)).digest('hex')).toBe('bc95c58ea9ff443697c4458ae733ed9a62eb53f436a948ee265cdd825a064742');expect(all.mappedDeclarations).toBe(494);expect(all.executionCredit).toBe(false);
});
test('tracking mapping omissions stale source unknown cases and altered direct assertions refuse',async()=>{
 const x=await input();for(const change of [(l:OutcomeMappingLedger)=>l.mappings.pop(),(l:OutcomeMappingLedger)=>{l.sourceSha256['src/docx/tracking-settings.ts']='0'.repeat(64);},(l:OutcomeMappingLedger)=>{l.mappings[0]!.caseKeys=['absent'];},(l:OutcomeMappingLedger)=>{l.mappings[0]!.assertions[0]='invented assertion';}]){const l=structuredClone(x.ledger);change(l);expect(()=>reconcileOutcomeMappings(x.inventory,l,x.features,x.sources)).toThrow();}
});
