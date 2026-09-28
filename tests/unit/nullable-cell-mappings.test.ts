import {beforeSharedPlannedExpansion} from './mapping-history.ts';
import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {inventoryNativeTests} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {outcomeMappingReport,reconcileOutcomeMappings,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
async function input(){const ledger=await Bun.file('docs/behaviors/nullable-cell-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};for(const p of Object.keys(ledger.sourceSha256))sources[p]=await Bun.file(p.startsWith('workflows/')?join(fixturesRoot(),p):p).text();return {ledger,sources,inventory:await inventoryNativeTests(),features:[parseFeature('workflows/docx/tables.feature',sources['workflows/docx/tables.feature']!,{allowDuplicateCaseNames:true})]};}
test('nullable cell mappings select one shared case with partial native associations and retain prior records except explicit inventory totals',async()=>{
 const x=await input(),r=reconcileOutcomeMappings(x.inventory,x.ledger,x.features,x.sources);expect(r.mappedDeclarations).toBe(7);expect(r.mappings.filter(m=>m.caseKeys.length).map(m=>m.caseKeys)).toEqual([['@id-docx-go-table-cell-access']]);expect(r.mappings.every(m=>m.status==='partial'&&m.gaps.length>0)).toBe(true);
 const all=await outcomeMappingReport(),old=structuredClone(all.mappings.filter(m=>!['nullable-cell','table-merging','template-inventory','comment-threads','revision-properties','revision-moves'].includes(m.ledger)));expect(old).toHaveLength(494);const adjusted=old.find(m=>m.testId==='bun:tests/unit/docx-heading-classification.test.ts:empty style retains strict Bun policy and all five shared getter-profile cases stay planned')!;expect(adjusted.assertions).toContain('expect(inv.counts.cases.implemented).toBe(732)');expect(adjusted.assertions).toContain('expect(inv.counts.cases.planned).toBe(59)');adjusted.assertions=adjusted.assertions.map(a=>a.replace('counts.cases.implemented).toBe(732)','counts.cases.implemented).toBe(550)').replace('counts.cases.planned).toBe(59)','counts.cases.planned).toBe(47)'));
 expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(beforeSharedPlannedExpansion(old))).digest('hex')).toBe('61f08e53ee37c013c54851a6429a2b3452234365a52b6b6979250961a906fda0');expect(all.mappedDeclarations).toBe(590);expect(all.executionCredit).toBe(false);
});
test('nullable mappings refuse stale sources missing declarations and unsupported case claims',async()=>{
 const x=await input();for(const change of [(l:OutcomeMappingLedger)=>l.mappings.pop(),(l:OutcomeMappingLedger)=>l.mappings[0]!.assertions.pop(),(l:OutcomeMappingLedger)=>{l.sourceSha256['src/docx/index.ts']='0'.repeat(64);},(l:OutcomeMappingLedger)=>{l.mappings[0]!.caseKeys=['unknown'];}]){const l=structuredClone(x.ledger);change(l);expect(()=>reconcileOutcomeMappings(x.inventory,l,x.features,x.sources)).toThrow();}
});
