import {test,expect} from 'bun:test';
import {outcomeMappingReport,reconcileOutcomeMappingSets,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
const paths=['tests/unit/xlsx-range.test.ts','tests/unit/xlsx-formula-analysis.test.ts','tests/unit/xlsx-formula-remap.test.ts'];
const canonical='workflows/xlsx/formula-references.feature';
async function sample(){const ledger=await Bun.file('docs/behaviors/formula-references-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};for(const p of Object.keys(ledger.sourceSha256))sources[p]=await Bun.file(p===canonical?join(fixturesRoot(),p):p).text();return {inventory:{cases:paths.flatMap(p=>inventoryTestSource(p,sources[p]!)),unresolved:[]},sets:[{name:'formula-references',expectedScopePaths:paths,ledger,sources,features:[parseFeature(canonical,sources[canonical]!,{allowDuplicateCaseNames:true})]}]};}

test('formula ledger enumerates all three native suites without changing execution credit or runtime leaf denominator',async()=>{
 const report=await outcomeMappingReport(),ledger=report.ledgers.find(l=>l.name==='formula-references');expect(ledger).toBeDefined();expect(ledger!.scopePaths).toEqual(paths);expect(ledger!.mappedDeclarations).toBe(31);expect(report.mappedDeclarations).toBe(446);expect(report.unmappedTestIds.length).toBe(report.totalDeclarations-446);expect(report.executionCredit).toBe(false);expect(report.runtimeLeafCount).toBeNull();const rows=report.mappings.filter(m=>m.ledger==='formula-references');expect(rows).toHaveLength(31);expect(rows.every(m=>m.status==='partial'&&m.executionCredit===false&&m.assertions.length&&m.outcomes.length&&m.gaps.length)).toBe(true);expect(new Set(rows.flatMap(m=>m.scenarioIds)).size).toBe(9);
});

test('formula mappings separate selected-case wrappers, incomplete analysis matrix and exact remap expressions',async()=>{
 const report=await outcomeMappingReport(),rows=report.mappings.filter(m=>m.ledger==='formula-references');
 const wrappers=rows.filter(m=>m.testId.includes('canonical'));expect(wrappers).toHaveLength(3);expect(wrappers.map(m=>m.caseKeys.length)).toEqual([13,19,13]);expect(wrappers.every(m=>m.gaps.join(' ').includes('aggregate status and counts'))).toBe(true);
 const analyzer=rows.find(m=>m.testId.includes('xlsx-formula-analysis.test.ts:finite expression matrix'))!;expect(analyzer.caseKeys).toEqual([]);expect(analyzer.gaps.join(' ')).toContain('does not call the remapper');
 const remap=rows.find(m=>m.testId.includes('xlsx-formula-remap.test.ts:finite 288-expression'))!;expect(remap.caseKeys).toEqual(['@id-xlsx-go-static-reference-properties']);expect(remap.gaps.join(' ')).toContain('not a fuzz campaign');
 const exact=rows.find(m=>m.testId.endsWith('static insertions rewrite exact shared expressions without touching literals or unrelated sheets'))!;expect(exact.caseKeys).toHaveLength(5);
 const limits=rows.filter(m=>m.testId.includes('output growth')||m.testId.includes('excessive input'));expect(limits.every(m=>m.caseKeys.length===0)).toBe(true);
});

test('formula ledger rejects omitted suites, declarations and exact assertions',async()=>{
 for(const p of paths){const a=await sample();a.sets[0]!.ledger.scopePaths=a.sets[0]!.ledger.scopePaths.filter(v=>v!==p);expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');const b=await sample();const index=b.sets[0]!.ledger.mappings.findIndex(m=>m.testId.includes(p));b.sets[0]!.ledger.mappings.splice(index,1);expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('Missing scoped');const c=await sample();c.sets[0]!.ledger.mappings.find(m=>m.testId.includes(p))!.assertions.pop();expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('assertion');}
});

test('formula source pins and canonical ownership refuse stale or fabricated evidence',async()=>{
 for(const p of ['src/xlsx/range.ts','src/xlsx/formula.ts','src/xlsx/formula-remap.ts','tests/acceptance/xlsx-range.ts','tests/acceptance/formula-analysis.ts','tests/acceptance/formula-remap.ts',canonical]){const a=await sample();a.sets[0]!.sources[p]+=' ';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('Stale source');}
 const a=await sample();a.sets[0]!.ledger.mappings[0]!.caseKeys=['@id-xlsx-go-static-reference-properties'];expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('case');const b=await sample();b.sets[0]!.ledger.mappings[0]!.scenarioIds=['@id-docx-go-new-empty-body'];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('scenario');const c=await sample();delete c.sets[0]!.ledger.sourceSha256['src/xlsx/formula-remap.ts'];expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('pin');
});

test('formula reconciliation retains source assertion arrays, review flags and input evidence unchanged',async()=>{
 const a=await sample(),before=JSON.stringify(a),r=reconcileOutcomeMappingSets(a.inventory,a.sets);expect(r.mappings).toHaveLength(31);expect(r.mappings.map(m=>m.assertions)).toEqual(a.inventory.cases.map(c=>c.assertions));expect(r.mappings.map(m=>m.reviewReasons)).toEqual(a.inventory.cases.map(c=>c.reviewReasons));expect(r.mappings.map(m=>m.deferredAssertions)).toEqual(a.inventory.cases.map(c=>c.deferredAssertions));expect(JSON.stringify(a)).toBe(before);expect(r.mappings.some(m=>m.reviewReasons.length>0)).toBe(true);
});

test('split literal and refusal predicates keep only their actually matched canonical case links',async()=>{
 const {inventory,sets}=await sample(),r=reconcileOutcomeMappingSets(inventory,sets),rows=r.mappings;
 const count=rows.find(m=>m.testId.includes('static formula analysis counts references'))!;expect(count.scenarioIds).toContain('@id-xlsx-go-formula-literal-punctuation');expect(count.caseKeys).toHaveLength(6);expect(count.caseKeys.every(k=>k.startsWith('@id-xlsx-go-formula-analysis-counts|'))).toBe(true);
 const bad=rows.find(m=>m.testId.includes('unsupported or incomplete grammar throws'))!;expect(bad.scenarioIds).toContain('@id-xlsx-go-formula-literal-punctuation');expect(bad.caseKeys).toHaveLength(7);expect(bad.caseKeys.every(k=>k.startsWith('@id-xlsx-go-formula-analysis-refusal|'))).toBe(true);
 const split=rows.filter(m=>m.testId.includes('xlsx-formula-remap.test.ts')&&m.scenarioIds.includes('@id-xlsx-go-static-remap-refusal')&&!m.testId.includes('canonical'));expect(split.length).toBeGreaterThan(0);expect(split.every(m=>m.caseKeys.length===0)).toBe(true);
 expect(rows.filter(m=>m.caseKeys.length)).toHaveLength(9);expect(rows.reduce((n,m)=>n+m.assertions.length,0)).toBe(127);
});
