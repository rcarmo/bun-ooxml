import {test,expect} from 'bun:test';
import {outcomeMappingReport,reconcileOutcomeMappingSets,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';import {fixturesRoot} from '../../scripts/fixture-inputs.ts';import {join} from 'node:path';
const paths=['tests/unit/docx-body-anchors.test.ts','tests/unit/docx-body-map.test.ts'];
const canonical='workflows/docx/anchor-discovery.feature',ids=['@id-python-word-anchor-headings-paragraphs','@id-python-word-anchor-text-filter','@id-python-word-anchor-document-map','@id-python-word-anchor-discover-insert'];
const pins=['src/docx/index.ts','src/docx/body-map.ts','src/docx/body-insertion.ts','src/docx/paragraph-properties.ts','tests/acceptance/body-anchors.ts','tests/acceptance/steps.ts','tests/fixtures/admission.ts'];
async function sample(){const ledger=await Bun.file('docs/behaviors/docx-anchors-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};for(const path of Object.keys(ledger.sourceSha256))sources[path]=await Bun.file(path===canonical?join(fixturesRoot(),path):path).text();return {inventory:{cases:paths.flatMap(p=>inventoryTestSource(p,sources[p]!)),unresolved:[]},sets:[{name:'docx-anchors',expectedScopePaths:paths,ledger,features:[parseFeature(canonical,sources[canonical]!,{allowDuplicateCaseNames:true})],sources}]};}

test('anchor mapping ledger accounts for 28 declarations and 122 direct assertions with prior mappings untouched',async()=>{
 const r=await outcomeMappingReport(),ledger=r.ledgers.find(l=>l.name==='docx-anchors');expect(ledger).toBeDefined();expect(ledger!.scopePaths).toEqual(paths);expect(ledger!.mappedDeclarations).toBe(28);const rows=r.mappings.filter(m=>m.ledger==='docx-anchors');expect(rows.reduce((n,m)=>n+m.assertions.length,0)).toBe(122);expect(rows.every(m=>m.status==='partial'&&m.executionCredit===false&&m.gaps.length&&m.outcomes.length)).toBe(true);expect(new Set(rows.flatMap(m=>m.scenarioIds))).toEqual(new Set(ids));expect(rows.filter(m=>m.caseKeys.length).map(m=>m.caseKeys.length)).toEqual([3,1]);expect(r.mappedDeclarations).toBe(546);expect(r.unmappedTestIds.length).toBe(r.totalDeclarations-546);expect(r.runtimeLeafCount).toBeNull();
 expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(r.mappings.filter(m=>['slide-order','effective-formatting','xml-values','pptx-core','docx-model','xlsx-comments','formula-references'].includes(m.ledger)))).digest('hex')).toBe('2e610f237d15c89c86027f9a75c60cc4e841a2f09ca9c14ee22e1b7d8127be3e');
});

test('anchor mappings retain saved-input differences, native map detail and actual negative-control attribution',async()=>{
 const rows=(await outcomeMappingReport()).mappings.filter(m=>m.ledger==='docx-anchors'),find=(suffix:string)=>rows.find(m=>m.testId.endsWith(suffix))!;
 const listing=find('body anchors read explicit outlines and paragraphs in saved order and filter case-insensitively without mutation');expect(listing.caseKeys).toEqual([]);expect(listing.gaps.join(' ')).toContain('byte reopen');expect(listing.gaps.join(' ')).toContain('Customer context');
 const insert=find('insert after captured heading saves immediately adjacent text and preserves every unrelated member');expect(insert.gaps.join(' ')).toContain('Current intro/Current delivery');expect(insert.gaps.join(' ')).toContain('bold is checked before save');
 const map=find('body map reports exact saved heading table placeholder and anchor counts without changing file or package bytes');expect(map.caseKeys).toEqual([]);expect(map.outcomes.join(' ')).toContain('real path');
 const exact=rows.filter(m=>m.caseKeys.length);expect(exact.every(m=>m.gaps.join(' ').includes('aggregate status and counts'))).toBe(true);expect(exact[0]!.gaps.join(' ')).toContain('two planned');expect(exact[1]!.gaps.join(' ')).toContain('four planned');
 const controls=find('canonical map predicates independently reject false counts empty anchors and changed file custody');expect(controls.gaps.join(' ')).toContain('AssertionError');expect(controls.gaps.join(' ')).toContain('nested');
 expect(rows.every(m=>!m.scenarioIds.some(id=>id.includes('hints')))).toBe(true);
});

test('anchor suites and all literal assertion expressions are mandatory and cannot borrow tool-hint cases',async()=>{
 for(const path of paths){const a=await sample();a.sets[0]!.ledger.scopePaths=a.sets[0]!.ledger.scopePaths.filter(p=>p!==path);expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');const b=await sample(),i=b.sets[0]!.ledger.mappings.findIndex(m=>m.testId.includes(path));expect(i).toBeGreaterThan(-1);b.sets[0]!.ledger.mappings.splice(i,1);expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('Missing scoped');const c=await sample();c.sets[0]!.ledger.mappings.find(m=>m.testId.includes(path))!.assertions.pop();expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('assertion');const d=await sample();d.sets[0]!.ledger.mappings.find(m=>m.testId.includes(path))!.caseKeys=['@id-python-word-anchor-section-discovery-hints'];expect(()=>reconcileOutcomeMappingSets(d.inventory,d.sets)).toThrow('case');}
});

test('anchor source pins independently reject stale runtime bindings helper and canonical files',async()=>{
 for(const path of [...paths,...pins,canonical]){const a=await sample();expect(Object.hasOwn(a.sets[0]!.ledger.sourceSha256,path)).toBe(true);a.sets[0]!.sources[path]+=' ';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('Stale source');const b=await sample();delete b.sets[0]!.ledger.sourceSha256[path];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('pin');}
});

test('anchor reconciliation preserves review flags nested assertions and the original evidence',async()=>{
 const a=await sample(),before=JSON.stringify(a),r=reconcileOutcomeMappingSets(a.inventory,a.sets);expect(JSON.stringify(a)).toBe(before);expect(r.mappings.map(m=>m.assertions)).toEqual(a.inventory.cases.map(m=>m.assertions));expect(r.mappings.map(m=>m.reviewReasons)).toEqual(a.inventory.cases.map(m=>m.reviewReasons));expect(r.mappings.map(m=>m.deferredAssertions)).toEqual(a.inventory.cases.map(m=>m.deferredAssertions));expect(r.executionCredit).toBe(false);expect(r.runtimeLeafCount).toBeNull();
});
