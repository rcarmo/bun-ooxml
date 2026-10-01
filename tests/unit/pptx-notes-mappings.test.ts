import {test,expect} from 'bun:test';
import {predecessorPptxFeature} from '../helpers/pptx-shared-predecessor.ts';
import {outcomeMappingReport,reconcileOutcomeMappingSets,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
const paths=['tests/unit/pptx.test.ts','tests/unit/pptx-notes-edit.test.ts','tests/unit/pptx-notes-utf8.test.ts'];
const pins=['src/pptx/notes.ts','tests/acceptance/notes-editing.ts','tests/acceptance/steps.ts','tests/fixtures/admission.ts'];
const canonical=["workflows/pptx/notes.feature","workflows/pptx/preservation.feature","workflows/pptx/text.feature"];
async function sample(){
 const ledger=await Bun.file('docs/behaviors/pptx-core-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};
 for(const path of new Set([...Object.keys(ledger.sourceSha256),...paths,...pins])){const text=await Bun.file(canonical.includes(path)?join(fixturesRoot(),path):path).text();sources[path]=canonical.includes(path)?await predecessorPptxFeature(path,text):text;}
 const inventory={cases:paths.flatMap(p=>inventoryTestSource(p,sources[p]!)),unresolved:[]};
 return {inventory,sets:[{name:'pptx-core',expectedScopePaths:paths,ledger,features:canonical.map(p=>parseFeature(p,sources[p]!,{allowDuplicateCaseNames:true})),sources}]};
}

test('PPTX notes mapping extension covers exactly two added suites while retaining all six prior records',async()=>{
 const report=await outcomeMappingReport(),ledger=report.ledgers.find(l=>l.name==='pptx-core')!;expect(ledger.scopePaths).toEqual(paths);expect(ledger.mappedDeclarations).toBe(33);
 const rows=report.mappings.filter(r=>r.ledger==='pptx-core').slice(6);expect(rows).toHaveLength(27);expect(rows.reduce((n,r)=>n+r.assertions.length,0)).toBe(109);expect(paths.slice(1).map(p=>rows.filter(r=>r.testId.includes(p)).length)).toEqual([14,13]);
 expect(rows.every(r=>r.status==='partial'&&r.executionCredit===false&&r.gaps.length&&r.outcomes.length)).toBe(true);expect(report.mappedDeclarations).toBe(590);expect(report.unmappedTestIds.length).toBe(report.totalDeclarations-590);expect(report.runtimeLeafCount).toBeNull();
 const raw=await Bun.file('docs/behaviors/pptx-core-mappings.json').json();expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(raw.mappings.slice(0,6))).digest('hex')).toBe('a6fda4a5457cafa44eba2181ffea1a2a374c42ede63f41df2b952cfb6c3b0f71');
 const exact=rows.filter(r=>r.caseKeys.length);expect(exact.map(r=>r.caseKeys.length)).toEqual([8,1]);expect(exact.every(r=>r.gaps.join(' ').includes('aggregate status and counts'))).toBe(true);
});

test('notes mappings keep byte/path, invalid-input, template and stale-source limits explicit',async()=>{
 const rows=(await outcomeMappingReport()).mappings.filter(r=>r.ledger==='pptx-core').slice(6),find=(suffix:string)=>rows.find(r=>r.testId.endsWith(suffix))!;
 const splice=find('existing notes replacement changes only its notes payload and reopens exact text');expect(splice.caseKeys).toEqual([]);expect(splice.gaps.join(' ')).toContain('byte reopen, not the required new path');
 const invalid=find('foreign, invalid and forged targets refuse atomically; exact no-op preserves bytes and target');expect(invalid.caseKeys).toEqual([]);expect(invalid.gaps.join(' ')).toContain('no raw FF input');expect(invalid.gaps.join(' ')).toContain('other session archive is not compared');
 const multiline=find('multiline uses only first-run template and preserves boundary empty paragraphs');expect(multiline.caseKeys).toEqual([]);expect(multiline.gaps.join(' ')).toContain('no direct four-paragraph count');
 const templates=find('multiline copies selected paragraph and first-run property fragments');expect(templates.gaps.join(' ')).toContain('b="1" differs from canonical b="0"');
 const saved=find('UTF8 notes replacement owns byte input, saves exact Unicode text and preserves every other member');expect(saved.caseKeys).toEqual([]);expect(saved.outcomes.join(' ')).toContain('real path save/reopen');expect(saved.gaps.join(' ')).toContain('does not compare an exact one-leaf lexical splice');
 const late=find('late protection and unsupported notes topology refuse byte no-ops without destructive repair');expect(late.gaps.join(' ')).toContain('stale-source checks may reject before topology or protection validation');
 const compound=find('canonical foreign invalid-byte and no-op case executes all four refusals and exact archive custody');expect(compound.gaps.join(' ')).toContain('string entry point');expect(compound.gaps.join(' ')).toContain('only the FF error code');
});

test('each added notes suite and its direct assertions cannot be omitted or borrow an unrelated case',async()=>{
 for(const path of paths.slice(1)){
  const a=await sample();a.sets[0]!.ledger.scopePaths=a.sets[0]!.ledger.scopePaths.filter(p=>p!==path);expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');
  const b=await sample(),index=b.sets[0]!.ledger.mappings.findIndex(r=>r.testId.includes(path));expect(index).toBeGreaterThan(-1);b.sets[0]!.ledger.mappings.splice(index,1);expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('Missing scoped');
  const c=await sample();c.sets[0]!.ledger.mappings.find(r=>r.testId.includes(path))!.assertions.pop();expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('assertion');
  const d=await sample();d.sets[0]!.ledger.mappings.find(r=>r.testId.includes(path))!.caseKeys=['@id-pptx-bun-open-save-noop'];expect(()=>reconcileOutcomeMappingSets(d.inventory,d.sets)).toThrow('case');
 }
});

test('notes runtime binding helper and native source pins reject independent stale or missing input',async()=>{
 for(const path of [...pins,...paths.slice(1)]){
  const a=await sample();expect(Object.hasOwn(a.sets[0]!.ledger.sourceSha256,path)).toBe(true);a.sets[0]!.sources[path]+=' ';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('Stale source');
  const b=await sample();delete b.sets[0]!.ledger.sourceSha256[path];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('pin');
 }
});

test('notes reconciliation retains dynamic and delegated assertion flags without altering inventory or mapping data',async()=>{
 const a=await sample(),before=JSON.stringify(a),report=reconcileOutcomeMappingSets(a.inventory,a.sets);expect(JSON.stringify(a)).toBe(before);
 expect(report.mappings.map(r=>r.assertions)).toEqual(a.inventory.cases.map(r=>r.assertions));expect(report.mappings.map(r=>r.reviewReasons)).toEqual(a.inventory.cases.map(r=>r.reviewReasons));expect(report.mappings.map(r=>r.deferredAssertions)).toEqual(a.inventory.cases.map(r=>r.deferredAssertions));expect(report.executionCredit).toBe(false);
 const row=report.mappings.find(r=>r.testId.endsWith('compound notes assertions detect each missing refusal, wrong decoding error and changed archive'))!;expect(row.caseKeys).toEqual([]);expect(row.gaps.join(' ')).toContain('Control-only');
});
