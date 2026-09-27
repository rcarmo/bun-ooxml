import {test,expect} from 'bun:test';
import {outcomeMappingReport,reconcileOutcomeMappingSets,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
const paths=['tests/unit/xlsx-comments.test.ts','tests/unit/xlsx-comment-vml-bindings.test.ts'];
const canonical='workflows/xlsx/comment-vml-custody.feature',id='@id-xlsx-comment-vml-existing-graph';
const pins=['src/xlsx/comments.ts','src/opc/package.ts','src/opc/content-types.ts','src/xml/index.ts','tests/acceptance/xlsx-comment-vml.ts','tests/acceptance/steps.ts','tests/fixtures/admission.ts','scripts/fixture-inputs.ts'];
async function sample(){
 const ledger=await Bun.file('docs/behaviors/xlsx-comments-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};
 for(const path of Object.keys(ledger.sourceSha256))sources[path]=await Bun.file(path===canonical?join(fixturesRoot(),path):path).text();
 return {inventory:{cases:paths.flatMap(p=>inventoryTestSource(p,sources[p]!)),unresolved:[]},sets:[{name:'xlsx-comments',expectedScopePaths:paths,ledger,features:[parseFeature(canonical,sources[canonical]!,{allowDuplicateCaseNames:true})],sources}]};
}

test('comment mapping ledger covers exactly 19 declarations and 72 direct expressions without mutation or execution credit',async()=>{
 const report=await outcomeMappingReport(),ledger=report.ledgers.find(r=>r.name==='xlsx-comments');expect(ledger).toBeDefined();expect(ledger!.scopePaths).toEqual(paths);expect(ledger!.mappedDeclarations).toBe(19);
 const rows=report.mappings.filter(r=>r.ledger==='xlsx-comments');expect(rows).toHaveLength(19);expect(rows.reduce((n,r)=>n+r.assertions.length,0)).toBe(72);expect(rows.every(r=>r.scenarioIds.length===1&&r.scenarioIds[0]===id)).toBe(true);
 expect(rows.every(r=>r.status==='partial'&&r.executionCredit===false&&r.gaps.length&&r.outcomes.length&&r.assertions.length)).toBe(true);
 expect(rows.filter(r=>r.caseKeys.length).map(r=>r.caseKeys)).toEqual([[id]]);expect(report.mappedDeclarations).toBe(501);expect(report.unmappedTestIds.length).toBe(report.totalDeclarations-501);expect(report.runtimeLeafCount).toBeNull();
});

test('comment mappings distinguish direct reads, delegated control counts, malformed graph custody and opaque VML',async()=>{
 const rows=(await outcomeMappingReport()).mappings.filter(r=>r.ledger==='xlsx-comments'),find=(suffix:string)=>rows.find(r=>r.testId.endsWith(suffix))!;
 const direct=find('native comment inspection returns detached exact authors refs text and separate VML relationship without writes');expect(direct.caseKeys).toEqual([]);expect(direct.gaps.join(' ')).toContain('no file save');
 const graph=find('external ambiguous mismatched and dangling comment or VML dependencies refuse unchanged');expect(graph.gaps.join(' ')).toContain('member arrays');
 const opaque=find('fresh inspections observe live metadata changes and protection does not turn reads into mutations');expect(opaque.gaps.join(' ')).toContain('opaque');expect(opaque.caseKeys).toEqual([]);
 const canonical=find('canonical XLSX comment and VML inspection checks both relationships and exact comment values');expect(canonical.gaps.join(' ')).toContain('aggregate status and counts');expect(canonical.gaps.join(' ')).toContain('five planned cases');
 const baseline=find('duplicate drawing IDs comment records and missing relationship dependencies fail independently of custody');expect(baseline.gaps.join(' ')).toContain('baseline serialization');expect(baseline.gaps.join(' ')).toContain('not necessarily the reader');
 const controls=rows.filter(r=>r.testId.includes('xlsx-comment-vml-bindings.test.ts')).filter(r=>!r.caseKeys.length);expect(controls).toHaveLength(6);expect(controls.every(r=>r.gaps.join(' ').includes('delegated'))).toBe(true);
});

test('comment suite declaration and exact direct assertion sets cannot silently shrink or borrow row-edit cases',async()=>{
 for(const path of paths){
  const a=await sample();a.sets[0]!.ledger.scopePaths=a.sets[0]!.ledger.scopePaths.filter(p=>p!==path);expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');
  const b=await sample(),index=b.sets[0]!.ledger.mappings.findIndex(r=>r.testId.includes(path));expect(index).toBeGreaterThan(-1);b.sets[0]!.ledger.mappings.splice(index,1);expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('Missing scoped');
  const c=await sample();c.sets[0]!.ledger.mappings.find(r=>r.testId.includes(path))!.assertions.pop();expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('assertion');
  const d=await sample();d.sets[0]!.ledger.mappings.find(r=>r.testId.includes(path))!.caseKeys=['@id-xlsx-comment-vml-disjoint-row-shift'];expect(()=>reconcileOutcomeMappingSets(d.inventory,d.sets)).toThrow('case');
 }
});

test('comment runtime binding helper and canonical pins are individually required and current',async()=>{
 for(const path of [...paths,...pins,canonical]){
  const a=await sample();expect(Object.hasOwn(a.sets[0]!.ledger.sourceSha256,path)).toBe(true);a.sets[0]!.sources[path]+=' ';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('Stale source');
  const b=await sample();delete b.sets[0]!.ledger.sourceSha256[path];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('pin');
 }
});

test('comment reconciliation retains original evidence, dynamic flags and nested assertion separation',async()=>{
 const a=await sample(),before=JSON.stringify(a),report=reconcileOutcomeMappingSets(a.inventory,a.sets);expect(JSON.stringify(a)).toBe(before);expect(report.mappings.map(r=>r.assertions)).toEqual(a.inventory.cases.map(r=>r.assertions));expect(report.mappings.map(r=>r.reviewReasons)).toEqual(a.inventory.cases.map(r=>r.reviewReasons));expect(report.mappings.map(r=>r.deferredAssertions)).toEqual(a.inventory.cases.map(r=>r.deferredAssertions));expect(report.executionCredit).toBe(false);
 expect(report.mappings.every(r=>!r.scenarioIds.some(id=>id.includes('row-')||id.includes('limited-editor')))).toBe(true);
});
