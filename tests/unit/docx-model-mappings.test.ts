import {test,expect} from 'bun:test';
import {outcomeMappingReport,reconcileOutcomeMappingSets,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
const paths=['tests/unit/docx-append-run.test.ts','tests/unit/docx-cell-properties.test.ts','tests/unit/docx-row-header.test.ts'];
const canonical='workflows/docx/document-model.feature';
async function sample(){
 const ledger=await Bun.file('docs/behaviors/docx-model-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};
 for(const path of Object.keys(ledger.sourceSha256))sources[path]=await Bun.file(path===canonical?join(fixturesRoot(),path):path).text();
 const inventory={cases:paths.flatMap(path=>inventoryTestSource(path,sources[path]!)),unresolved:[]};
 const features=[parseFeature(canonical,sources[canonical]!,{allowDuplicateCaseNames:true})];
 return {inventory,sets:[{name:'docx-model',expectedScopePaths:paths,ledger,features,sources}]};
}

test('committed DOCX model ledger covers exactly three full native suites without new execution credit',async()=>{
 const report=await outcomeMappingReport(),ledger=report.ledgers.find(l=>l.name==='docx-model');
 expect(ledger).toBeDefined();expect(ledger!.scopePaths).toEqual(paths);expect(ledger!.mappedDeclarations).toBe(38);
 const mappings=report.mappings.filter(m=>m.ledger==='docx-model');expect(mappings).toHaveLength(38);
 expect(mappings.every(m=>m.status==='partial'&&m.executionCredit===false&&m.gaps.length&&m.outcomes.length&&m.assertions.length)).toBe(true);
 expect(report.executionCredit).toBe(false);expect(report.runtimeLeafCount).toBeNull();expect(report.mappedDeclarations).toBe(82);expect(report.unmappedTestIds.length).toBe(report.totalDeclarations-82);
 expect(new Set(mappings.flatMap(m=>m.scenarioIds))).toEqual(new Set(['@id-docx-go-paragraph-multiple-runs','@id-docx-go-roundtrip-selected-formatting','@id-docx-go-cell-shading-getter','@id-docx-go-cell-properties-getters','@id-docx-go-table-header-getter']));
});

test('DOCX mappings distinguish aggregate acceptance assertions, native extensions and different table dimensions',async()=>{
 const report=await outcomeMappingReport(),rows=report.mappings.filter(m=>m.ledger==='docx-model');
 const aggregates=rows.filter(m=>m.testId.includes('shared'));
 expect(aggregates).toHaveLength(3);expect(aggregates.every(m=>m.gaps.join(' ').includes('aggregate status and counts'))).toBe(true);
 expect(aggregates.map(m=>m.caseKeys.length)).toEqual([2,2,1]);
 const rowToggle=rows.find(m=>m.testId.endsWith('direct row header toggles preserve text, dimensions and other members after path save/reopen'))!;
 expect(rowToggle.caseKeys).toEqual([]);expect(rowToggle.gaps.join(' ')).toContain('2x2');expect(rowToggle.gaps.join(' ')).toContain('3x2');
 expect(rowToggle.outcomes.join(' ')).toContain('live dimensions and Header/Body text are checked before save');
 expect(rowToggle.outcomes.join(' ')).toContain('Real path reopen retains true/false row states, the Body cell text and unrelated members');
 const cell=rows.find(m=>m.testId.endsWith('direct cell properties survive reopen without changing text, grid, sibling cells or other package parts'))!;
 expect(cell.outcomes.join(' ')).toContain('live dimensions and selected text are checked before save');
 expect(cell.outcomes.join(' ')).toContain('After reopen, exact sibling-cell XML and every unrelated member match');
 const outputLimits=rows.filter(m=>m.testId.includes('output bound'));
 expect(outputLimits).toHaveLength(2);expect(outputLimits.every(m=>m.caseKeys.length===0&&m.gaps.join(' ').includes('Control-only'))).toBe(true);
 const formatting=rows.find(m=>m.testId.endsWith('per-run authoring preserves selected formatting after real path save and reopen'))!;
 expect(formatting.caseKeys).toEqual(['@id-docx-go-roundtrip-selected-formatting']);expect(formatting.gaps.join(' ')).toContain('No unrelated-part');
 const utf16=rows.filter(m=>m.testId.includes('UTF-16'));expect(utf16).toHaveLength(3);expect(utf16.every(m=>m.caseKeys.length===0&&m.gaps.join(' ').includes('UTF-16'))).toBe(true);
});

test('DOCX ledger refuses dropped suites, dropped declarations, omitted assertions and invented assertions',async()=>{
 const a=await sample();a.sets[0]!.ledger.scopePaths.pop();expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');
 const b=await sample();b.sets[0]!.ledger.mappings.pop();expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('Missing scoped');
 const c=await sample();c.sets[0]!.ledger.mappings[0]!.assertions.pop();expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('assertion');
 const d=await sample();d.sets[0]!.ledger.mappings[0]!.assertions[0]='expect(fake).toBe(true)';expect(()=>reconcileOutcomeMappingSets(d.inventory,d.sets)).toThrow('assertion');
});

test('DOCX ledger refuses stale runtime, binding and canonical pins and unrelated canonical identities',async()=>{
 for(const path of ['src/docx/append-run.ts','src/docx/cell-properties.ts','src/docx/row-header.ts','src/docx/index.ts','tests/acceptance/docx-model.ts',canonical]){
  const a=await sample();a.sets[0]!.sources[path]+=' ';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('Stale source');
 }
 const b=await sample();b.sets[0]!.ledger.mappings[0]!.scenarioIds=['@id-xml-parse-simple'];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('scenario');
 const c=await sample();c.sets[0]!.ledger.mappings[0]!.caseKeys=['@id-docx-go-table-header-getter'];expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('case');
 const d=await sample();delete d.sets[0]!.ledger.sourceSha256['src/docx/row-header.ts'];expect(()=>reconcileOutcomeMappingSets(d.inventory,d.sets)).toThrow('pin');
});

test('DOCX mapping validation retains native dynamic review flags and leaves source evidence untouched',async()=>{
 const a=await sample(),before=JSON.stringify(a),report=reconcileOutcomeMappingSets(a.inventory,a.sets);
 expect(report.mappings).toHaveLength(38);expect(JSON.stringify(a)).toBe(before);
 expect(report.mappings.map(m=>m.assertions)).toEqual(a.inventory.cases.map(c=>c.assertions));
 expect(report.mappings.map(m=>m.reviewReasons)).toEqual(a.inventory.cases.map(c=>c.reviewReasons));
 expect(report.mappings.map(m=>m.deferredAssertions)).toEqual(a.inventory.cases.map(c=>c.deferredAssertions));
 expect(report.mappings.some(m=>m.reviewReasons.length>0)).toBe(true);
 const ownCases=report.mappings.filter(m=>m.caseKeys.length>0);expect(ownCases).toHaveLength(5);
});
