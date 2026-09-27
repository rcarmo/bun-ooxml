import {test,expect} from 'bun:test';
import {outcomeMappingReport,reconcileOutcomeMappingSets,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
const paths=['tests/unit/docx-append-run.test.ts','tests/unit/docx-cell-properties.test.ts','tests/unit/docx-row-header.test.ts','tests/unit/docx-paragraph-text.test.ts','tests/unit/docx-insert-paragraph.test.ts','tests/unit/docx-table-rows.test.ts','tests/unit/docx-row-texts.test.ts','tests/unit/docx-core-properties.test.ts','tests/unit/docx-document-properties.test.ts','tests/unit/docx-table-style.test.ts'];
const canonical=["workflows/docx/page-layout.feature","workflows/docx/paragraphs.feature","workflows/docx/properties.feature","workflows/docx/run-formatting.feature","workflows/docx/tables.feature"];
async function sample(){
 const ledger=await Bun.file('docs/behaviors/docx-model-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};
 for(const path of Object.keys(ledger.sourceSha256))sources[path]=await Bun.file(canonical.includes(path)?join(fixturesRoot(),path):path).text();
 const inventory={cases:paths.flatMap(path=>inventoryTestSource(path,sources[path]!)),unresolved:[]};
 const features=canonical.map(p=>parseFeature(p,sources[p]!,{allowDuplicateCaseNames:true}));
 return {inventory,sets:[{name:'docx-model',expectedScopePaths:paths,ledger,features,sources}]};
}

test('committed DOCX model ledger covers exactly ten full native suites without new execution credit',async()=>{
 const report=await outcomeMappingReport(),ledger=report.ledgers.find(l=>l.name==='docx-model');
 expect(ledger).toBeDefined();expect(ledger!.scopePaths).toEqual(paths);expect(ledger!.mappedDeclarations).toBe(130);
 const mappings=report.mappings.filter(m=>m.ledger==='docx-model');expect(mappings).toHaveLength(130);
 expect(mappings.every(m=>m.status==='partial'&&m.executionCredit===false&&m.gaps.length&&m.outcomes.length&&m.assertions.length)).toBe(true);
 expect(report.executionCredit).toBe(false);expect(report.runtimeLeafCount).toBeNull();expect(report.mappedDeclarations).toBe(494);expect(report.unmappedTestIds.length).toBe(report.totalDeclarations-494);
 expect(new Set(mappings.flatMap(m=>m.scenarioIds))).toEqual(new Set(['@id-docx-go-paragraph-multiple-runs','@id-docx-go-roundtrip-selected-formatting','@id-docx-go-cell-shading-getter','@id-docx-go-cell-properties-getters','@id-docx-go-table-header-getter','@id-docx-go-paragraph-text-getter','@id-docx-go-body-insert-order','@id-docx-go-table-row-counts','@id-docx-go-table-cell-text-getters','@id-docx-go-core-properties-getters','@id-docx-go-section-title-background-getters','@id-docx-go-table-style-getter']));
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
 expect(outputLimits).toHaveLength(8);expect(outputLimits.every(m=>m.caseKeys.length===0&&m.gaps.join(' ').includes('Control-only'))).toBe(true);
 const formatting=rows.find(m=>m.testId.endsWith('per-run authoring preserves selected formatting after real path save and reopen'))!;
 expect(formatting.caseKeys).toEqual(['@id-docx-go-roundtrip-selected-formatting']);expect(formatting.gaps.join(' ')).toContain('No unrelated-part');
 const utf16=rows.filter(m=>m.testId.includes('UTF-16'));expect(utf16).toHaveLength(7);expect(utf16.every(m=>m.caseKeys.length===0&&m.gaps.join(' ').includes('UTF-16'))).toBe(true);
});

test('DOCX ledger refuses dropped suites, dropped declarations, omitted assertions and invented assertions',async()=>{
 const a=await sample();a.sets[0]!.ledger.scopePaths.pop();expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');
 const b=await sample();b.sets[0]!.ledger.mappings.pop();expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('Missing scoped');
 const c=await sample();c.sets[0]!.ledger.mappings[0]!.assertions.pop();expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('assertion');
 const d=await sample();d.sets[0]!.ledger.mappings[0]!.assertions[0]='expect(fake).toBe(true)';expect(()=>reconcileOutcomeMappingSets(d.inventory,d.sets)).toThrow('assertion');
});

test('DOCX ledger refuses stale runtime, binding and canonical pins and unrelated canonical identities',async()=>{
 for(const path of ['src/docx/append-run.ts','src/docx/cell-properties.ts','src/docx/row-header.ts','src/docx/index.ts','tests/acceptance/docx-model.ts',...canonical]){
  const a=await sample();a.sets[0]!.sources[path]+=' ';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('Stale source');
 }
 const b=await sample();b.sets[0]!.ledger.mappings[0]!.scenarioIds=['@id-xml-parse-simple'];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('scenario');
 const c=await sample();c.sets[0]!.ledger.mappings[0]!.caseKeys=['@id-docx-go-table-header-getter'];expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('case');
 const d=await sample();delete d.sets[0]!.ledger.sourceSha256['src/docx/row-header.ts'];expect(()=>reconcileOutcomeMappingSets(d.inventory,d.sets)).toThrow('pin');
});

test('DOCX mapping validation retains native dynamic review flags and leaves source evidence untouched',async()=>{
 const a=await sample(),before=JSON.stringify(a),report=reconcileOutcomeMappingSets(a.inventory,a.sets);
 expect(report.mappings).toHaveLength(130);expect(JSON.stringify(a)).toBe(before);
 expect(report.mappings.map(m=>m.assertions)).toEqual(a.inventory.cases.map(c=>c.assertions));
 expect(report.mappings.map(m=>m.reviewReasons)).toEqual(a.inventory.cases.map(c=>c.reviewReasons));
 expect(report.mappings.map(m=>m.deferredAssertions)).toEqual(a.inventory.cases.map(c=>c.deferredAssertions));
 expect(report.mappings.some(m=>m.reviewReasons.length>0)).toBe(true);
 const ownCases=report.mappings.filter(m=>m.caseKeys.length>0);expect(ownCases).toHaveLength(12);
});

test('expanded DOCX mappings retain input, save timing and delegated predicate gaps',async()=>{
 const report=await outcomeMappingReport(),rows=report.mappings.filter(m=>m.ledger==='docx-model').slice(38,88);
 expect(rows).toHaveLength(50);expect(rows.reduce((n,m)=>n+m.assertions.length,0)).toBe(262);
 const groups=['docx-paragraph-text.test.ts','docx-insert-paragraph.test.ts','docx-table-rows.test.ts','docx-row-texts.test.ts'];
 expect(groups.map(path=>rows.filter(m=>m.testId.includes(path)).length)).toEqual([13,13,14,10]);
 const exact=rows.filter(m=>m.caseKeys.length);expect(exact).toHaveLength(4);expect(exact.map(m=>m.caseKeys.length)).toEqual([5,1,1,1]);expect(exact.every(m=>m.gaps.join(' ').includes('aggregate status and counts'))).toBe(true);
 const sequential=rows.find(m=>m.testId.endsWith('whole paragraph text replaces empty and existing values with fresh handles and saved readback'))!;
 expect(sequential.caseKeys).toEqual([]);expect(sequential.gaps.join(' ')).toContain('sequentially');expect(sequential.gaps.join(' ')).toContain('five outline values use byte reopen');
 const rowCounts=rows.find(m=>m.testId.endsWith('append insert delete return fresh tables and retain row order after real path save'))!;
 expect(rowCounts.caseKeys).toEqual([]);expect(rowCounts.gaps.join(' ')).toContain('2x3, not local 2x2');expect(rowCounts.gaps.join(' ')).toContain('Only the final state is path-reopened');
 const firstRow=rows.find(m=>m.testId.endsWith('row texts read exact cells in column order and detach results without archive mutation'))!;
 expect(firstRow.gaps.join(' ')).toContain('only cell(0,0), not all four');expect(firstRow.caseKeys).toEqual([]);
 expect(rows.every(m=>m.executionCredit===false&&m.status==='partial')).toBe(true);
});

test('each added DOCX suite and every new runtime/binding pin is required independently',async()=>{
 for(const path of paths.slice(3)){
  const a=await sample();a.sets[0]!.ledger.scopePaths=a.sets[0]!.ledger.scopePaths.filter(p=>p!==path);expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');
  const b=await sample();const index=b.sets[0]!.ledger.mappings.findIndex(m=>m.testId.includes(path));expect(index).toBeGreaterThan(-1);b.sets[0]!.ledger.mappings.splice(index,1);expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('Missing scoped');
 }
 for(const path of ['tests/acceptance/paragraph-text.ts','tests/acceptance/paragraph-properties.ts','tests/acceptance/body-insertion.ts','tests/acceptance/table-rows.ts','tests/acceptance/row-texts.ts','src/docx/paragraph-text.ts','src/docx/body-insertion.ts','src/docx/table-rows.ts']){
  const a=await sample();expect(Object.hasOwn(a.sets[0]!.ledger.sourceSha256,path)).toBe(true);a.sets[0]!.sources[path]+=' ';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('Stale source');
 }
});

test('new mappings cannot borrow a sibling model case or suppress assertion expressions',async()=>{
 for(const path of paths.slice(3)){
  const a=await sample(),row=a.sets[0]!.ledger.mappings.find(m=>m.testId.includes(path))!;row.assertions.pop();expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('assertion');
  const b=await sample(),other=b.sets[0]!.ledger.mappings.find(m=>m.testId.includes(path))!;other.caseKeys=['@id-docx-go-table-header-getter'];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('case');
 }
 const a=await sample(),row=a.sets[0]!.ledger.mappings.find(m=>m.testId.includes('five canonical text getter cases'))!;row.caseKeys=['@id-docx-go-paragraph-text-getter|invented'];expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('case');
});
