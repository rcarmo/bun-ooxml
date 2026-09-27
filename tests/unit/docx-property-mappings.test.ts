import {test,expect} from 'bun:test';
import {outcomeMappingReport,reconcileOutcomeMappingSets,type OutcomeMappingLedger} from '../../scripts/outcome-mappings.ts';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
const paths=['tests/unit/docx-core-properties.test.ts','tests/unit/docx-document-properties.test.ts','tests/unit/docx-table-style.test.ts'];
const ids=['@id-docx-go-core-properties-getters','@id-docx-go-section-title-background-getters','@id-docx-go-table-style-getter'];
const pins=['src/opc/core-properties.ts','src/docx/document-properties.ts','src/docx/table-style.ts','tests/acceptance/core-properties.ts','tests/acceptance/document-properties.ts','tests/acceptance/table-style.ts'];
const canonical=["workflows/docx/page-layout.feature","workflows/docx/paragraphs.feature","workflows/docx/properties.feature","workflows/docx/run-formatting.feature","workflows/docx/tables.feature"];
async function sample(){
 const ledger=await Bun.file('docs/behaviors/docx-model-mappings.json').json() as OutcomeMappingLedger,sources:Record<string,string>={};
 for(const path of Object.keys(ledger.sourceSha256))sources[path]=await Bun.file(canonical.includes(path)?join(fixturesRoot(),path):path).text();
 const inventory={cases:ledger.scopePaths.flatMap(path=>inventoryTestSource(path,sources[path]!)),unresolved:[]};
 return {inventory,sets:[{name:'docx-model',expectedScopePaths:ledger.scopePaths.slice(),ledger,features:canonical.map(p=>parseFeature(p,sources[p]!,{allowDuplicateCaseNames:true})),sources}]};
}

test('property mapping extension adds exactly 42 declarations and 226 direct expressions while retaining prior records',async()=>{
 const report=await outcomeMappingReport(),all=report.mappings.filter(m=>m.ledger==='docx-model'),rows=all.slice(88);
 expect(rows).toHaveLength(42);expect(rows.reduce((n,r)=>n+r.assertions.length,0)).toBe(226);
 expect(paths.map(p=>rows.filter(r=>r.testId.startsWith('bun:'+p+':')).length)).toEqual([15,15,12]);
 expect(rows.every(r=>r.status==='partial'&&r.executionCredit===false&&r.outcomes.length&&r.gaps.length)).toBe(true);
 expect(new Set(rows.flatMap(r=>r.scenarioIds))).toEqual(new Set(ids));
 const exact=rows.filter(r=>r.caseKeys.length);expect(exact.map(r=>r.caseKeys)).toEqual(ids.map(id=>[id]));
 expect(exact.every(r=>r.testId.includes(':canonical ')&&r.gaps.join(' ').includes('aggregate status and counts'))).toBe(true);
 const ledger=await Bun.file('docs/behaviors/docx-model-mappings.json').json();
 expect(new Bun.CryptoHasher('sha256').update(JSON.stringify(ledger.mappings.slice(0,88))).digest('hex')).toBe('34cfb0ba78d652e98d0dc50afac5241d4265b3e59cb036b4978334d0143e0dc1');
 expect(report.mappedDeclarations).toBe(590);expect(report.unmappedTestIds.length).toBe(report.totalDeclarations-590);
});

test('property mappings distinguish selected fields, live custody, reopened checks and unresolved style references',async()=>{
 const report=await outcomeMappingReport(),rows=report.mappings.filter(m=>m.ledger==='docx-model').slice(88),find=(name:string)=>rows.find(r=>r.testId.endsWith(name))!;
 const core=find('all fifteen core fields survive real path save/reopen with exact unrelated payload custody');
 expect(core.caseKeys).toEqual([]);expect(core.gaps.join(' ')).toContain('only title, creator and subject');expect(core.gaps.join(' ')).toContain('relationship and MIME checks occur before save');
 const doc=find('direct title-page and background values persist through path reopen without changing content or unrelated members');
 expect(doc.caseKeys).toEqual([]);expect(doc.gaps.join(' ')).toContain('Table dimensions and stale handles are checked before save');
 const style=find('unresolved direct table style survives path reopen without registry creation or unrelated changes');
 expect(style.caseKeys).toEqual([]);expect(style.gaps.join(' ')).toContain('undefined');expect(style.gaps.join(' ')).toContain('no style resolution or rendered appearance');
 expect(rows.filter(r=>r.testId.includes('output bound')).map(r=>r.caseKeys)).toEqual([[],[],[]]);
 const wrappers=rows.filter(r=>r.caseKeys.length);expect(wrappers.every(r=>r.gaps.join(' ').includes('No save/reopen'))).toBe(true);
 expect(wrappers[0]!.gaps.join(' ')).toContain('cached getter result');expect(wrappers[2]!.gaps.join(' ')).toContain('cached before/after strings');
});

test('each property suite, declaration and assertion remains mandatory and cannot borrow a sibling canonical case',async()=>{
 for(const [i,path]of paths.entries()){
  const a=await sample();a.sets[0]!.ledger.scopePaths=a.sets[0]!.ledger.scopePaths.filter(p=>p!==path);expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');
  const b=await sample(),index=b.sets[0]!.ledger.mappings.findIndex(r=>r.testId.startsWith('bun:'+path+':'));expect(index).toBeGreaterThan(-1);b.sets[0]!.ledger.mappings.splice(index,1);expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('Missing scoped');
  const c=await sample();c.sets[0]!.ledger.mappings.find(r=>r.testId.startsWith('bun:'+path+':'))!.assertions.pop();expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('assertion');
  const d=await sample();d.sets[0]!.ledger.mappings.find(r=>r.testId.startsWith('bun:'+path+':'))!.caseKeys=[ids[(i+1)%ids.length]!];expect(()=>reconcileOutcomeMappingSets(d.inventory,d.sets)).toThrow('case');
 }
});

test('property runtime, binding and native pins independently refuse missing and stale sources',async()=>{
 for(const path of [...pins,...paths]){
  const a=await sample();expect(Object.hasOwn(a.sets[0]!.ledger.sourceSha256,path)).toBe(true);a.sets[0]!.sources[path]+=' ';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('Stale source');
  const b=await sample();delete b.sets[0]!.ledger.sourceSha256[path];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('pin');
 }
});
