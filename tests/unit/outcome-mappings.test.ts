import {test,expect} from 'bun:test';
import {createHash} from 'node:crypto';
import {inventoryTestSource} from '../../scripts/test-inventory.ts';
import {parseFeature} from '../../scripts/gherkin.ts';
import {reconcileOutcomeMappings,reconcileOutcomeMappingSets,type OutcomeMappingLedger,type OutcomeMappingSet} from '../../scripts/outcome-mappings.ts';
const source=`import {test,expect} from 'bun:test';test('saved outcome',()=>{expect(1).toBe(1)});test('unmapped',()=>{expect(2).toBe(2)});`;
const feature=`@planned
Feature: Saved outcome
 @id-pptx-mapping-test
 Scenario Outline: Saved <kind>
  Given input <kind>
  Then saved value matches
  Examples:
   | kind |
   | plain |
   | alias |
`;
function sample(){
 const records=inventoryTestSource('tests/unit/sample.test.ts',source);
 const inventory={cases:records,unresolved:[]};
 const features=[parseFeature('workflows/pptx/sample.feature',feature)];
 const sources={'tests/unit/sample.test.ts':source,'workflows/pptx/sample.feature':feature};
 const ledger:OutcomeMappingLedger={schemaVersion:1,consumer:'bun',executionCredit:false,scopePaths:['tests/unit/sample.test.ts'],sourceSha256:Object.fromEntries(Object.entries(sources).map(([p,s])=>[p,createHash('sha256').update(s).digest('hex')])),mappings:records.map(c=>({testId:c.id,status:'partial',scenarioIds:['@id-pptx-mapping-test'],caseKeys:[],assertions:[...c.assertions],outcomes:['A concrete saved-value assertion'],gaps:['Only this bounded assertion is mapped']}))};
 return {inventory,features,sources,ledger};
}
test('outcome mappings validate source pins and preserve the unmapped denominator without execution credit',()=>{
 const {inventory,features,sources,ledger}=sample();
 inventory.cases.push({...inventory.cases[0]!,id:'bun:elsewhere:unmapped',path:'tests/unit/other.test.ts'});
 const r=reconcileOutcomeMappings(inventory,ledger,features,sources) as any;
 expect(r.mappedDeclarations).toBe(2);expect(r.totalDeclarations).toBe(3);expect(r.unmappedTestIds).toEqual(['bun:elsewhere:unmapped']);expect(r.executionCredit).toBe(false);expect(r.mappings.every((m:any)=>m.executionCredit===false&&m.status==='partial')).toBe(true);
});
test('mapping gate refuses missing or duplicated test declarations',()=>{
 const {inventory,features,sources,ledger}=sample();ledger.mappings.pop();expect(()=>reconcileOutcomeMappings(inventory,ledger,features,sources)).toThrow('Missing');
 ledger.mappings.push(ledger.mappings[0]!);expect(()=>reconcileOutcomeMappings(inventory,ledger,features,sources)).toThrow('Duplicate');
});
test('mapping gate refuses stale source or nonexistent native assertions',()=>{
 const a=sample();a.sources['tests/unit/sample.test.ts']+=' ';expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('source');
 const b=sample();b.ledger.mappings[0]!.assertions=['expect(1).toBe(99)'];expect(()=>reconcileOutcomeMappings(b.inventory,b.ledger,b.features,b.sources)).toThrow('assertion');
});
test('mapping gate refuses unknown canonical scenarios, case keys and duplicate IDs',()=>{
 const a=sample();a.ledger.mappings[0]!.scenarioIds=['@id-absent'];expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('scenario');
 const b=sample();b.ledger.mappings[0]!.caseKeys=['@id-pptx-mapping-test|{"kind":"absent"}'];expect(()=>reconcileOutcomeMappings(b.inventory,b.ledger,b.features,b.sources)).toThrow('case');
 const c=sample();c.features.push(c.features[0]!);expect(()=>reconcileOutcomeMappings(c.inventory,c.ledger,c.features,c.sources)).toThrow('Duplicate');
});
test('mapping gate rejects missing source pins, invented scope paths and empty outcomes/gaps',()=>{
 const a=sample();delete a.ledger.sourceSha256['workflows/pptx/sample.feature'];expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('pin');
 const b=sample();b.ledger.scopePaths.push('tests/unit/missing.test.ts');expect(()=>reconcileOutcomeMappings(b.inventory,b.ledger,b.features,b.sources)).toThrow('scope');
 const c=sample();c.ledger.mappings[0]!.outcomes=[];expect(()=>reconcileOutcomeMappings(c.inventory,c.ledger,c.features,c.sources)).toThrow('outcome');
 const d=sample();d.ledger.mappings[0]!.gaps=[];expect(()=>reconcileOutcomeMappings(d.inventory,d.ledger,d.features,d.sources)).toThrow('gap');
});
test('mapping gate rejects fabricated execution credit and preserves dynamic review flags',()=>{
 const a=sample();(a.ledger as any).executionCredit=true;expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('credit');
 const b=sample();b.inventory.cases[0]!.reviewReasons.push('body-loop');
 const r=reconcileOutcomeMappings(b.inventory,b.ledger,b.features,b.sources) as any;
 expect(r.mappings[0].reviewReasons).toEqual(['body-loop']);expect(r.runtimeLeafCount).toBe(null);
});

test('committed ledgers enumerate slide-order, effective-formatting, XML, PPTX and DOCX model assertions with known gaps',async()=>{
 const {outcomeMappingReport}=await import('../../scripts/outcome-mappings.ts');
 const report=await outcomeMappingReport();
 expect(report.mappedDeclarations).toBe(373);
 expect(report.unmappedTestIds.length).toBe(report.totalDeclarations-373);
 expect(report.ledgers.map(l=>[l.name,l.mappedDeclarations])).toEqual([['slide-order',13],['effective-formatting',12],['xml-values',13],['pptx-core',33],['docx-model',130],['docx-anchors',28],['xlsx-comments',19],['xlsx-styles',16],['pptx-text-boxes',11],['docx-page-layout',13],['docx-paragraph-styles',17],['docx-style-authoring',11],['xml-comparison',14],['package-comparison',12],['formula-references',31]]);
 expect(report.mappings[0]!.caseKeys).toHaveLength(21);
 expect(report.mappings.find(m=>m.ledger==='effective-formatting')!.caseKeys).toHaveLength(25);
 expect(report.mappings.every(m=>m.gaps.length>0&&m.outcomes.length>0&&m.assertions.length>0&&m.executionCredit===false)).toBe(true);
});

test('mapping gate refuses silently omitted native assertions',()=>{
 const a=sample();a.inventory.cases[0]!.assertions.push('expect(3).toBe(3)');
 expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('assertion');
});
test('scenario-only links remain explicitly distinct from case-key links',()=>{
 const a=sample();a.ledger.mappings[1]!.caseKeys=[a.features[0]!.scenarios[0]!.cases[0]!.identityKey];
 const r=reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources) as any;
 expect(r.mappings[0].linkGranularity).toBe('scenario-only');
 expect(r.mappings[1].linkGranularity).toBe('explicit-case-keys');
 expect(r.mappings.every((m:any)=>m.executionCredit===false&&m.status==='partial')).toBe(true);
});

function twoSets(){
 const a=sample(),b=sample(),oldPath='tests/unit/sample.test.ts',path='tests/unit/second.test.ts',featurePath='workflows/docx/second.feature',secondFeature=feature.replaceAll('@id-pptx-mapping-test','@id-docx-second-test');
 b.inventory.cases=inventoryTestSource(path,source);
 b.features=[parseFeature(featurePath,secondFeature)];
 const sources:Record<string,string>={[path]:source,[featurePath]:secondFeature};
 b.ledger.scopePaths=[path];b.ledger.sourceSha256=Object.fromEntries(Object.entries(sources).map(([p,s])=>[p,createHash('sha256').update(s).digest('hex')]));
 b.ledger.mappings.forEach((m,i)=>{m.testId=b.inventory.cases[i]!.id;m.scenarioIds=['@id-docx-second-test'];});
 const sets:OutcomeMappingSet[]=[{name:'first',expectedScopePaths:[oldPath],ledger:a.ledger,features:a.features,sources:a.sources},{name:'second',expectedScopePaths:[path],ledger:b.ledger,features:b.features,sources}];
 return {inventory:{cases:[...a.inventory.cases,...b.inventory.cases],unresolved:[]},sets};
}
test('independent ledgers share one denominator without execution credit or dropped mappings',()=>{
 const {inventory,sets}=twoSets();inventory.cases.push({...inventory.cases[0]!,id:'bun:outside',path:'tests/unit/outside.test.ts'});
 const r=reconcileOutcomeMappingSets(inventory,sets) as any;
 expect(r.mappedDeclarations).toBe(4);expect(r.totalDeclarations).toBe(5);expect(r.unmappedTestIds).toEqual(['bun:outside']);
 expect(r.ledgers.map((l:any)=>[l.name,l.mappedDeclarations])).toEqual([['first',2],['second',2]]);
 expect(r.mappings.map((m:any)=>m.ledger)).toEqual(['first','first','second','second']);expect(r.executionCredit).toBe(false);expect(r.runtimeLeafCount).toBe(null);
});
test('independent ledgers cannot borrow sibling scenarios or source pins',()=>{
 const a=twoSets();a.sets[0]!.ledger.mappings[0]!.scenarioIds=['@id-docx-second-test'];expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scenario');
 const b=twoSets();delete b.sets[0]!.ledger.sourceSha256['workflows/pptx/sample.feature'];expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('pin');
});
test('independent ledgers reject duplicate names, scopes and canonical scenarios',()=>{
 const a=twoSets();a.sets[1]!.name='first';expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('name');
 const b=twoSets();b.sets[1]={...b.sets[0]!,name:'second'};expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('scope');
 const c=twoSets();c.sets[1]!.features=c.sets[0]!.features;c.sets[1]!.sources['workflows/pptx/sample.feature']=feature;c.sets[1]!.ledger.sourceSha256['workflows/pptx/sample.feature']=createHash('sha256').update(feature).digest('hex');c.sets[1]!.ledger.mappings.forEach(m=>m.scenarioIds=['@id-pptx-mapping-test']);expect(()=>reconcileOutcomeMappingSets(c.inventory,c.sets)).toThrow('scenario');
});
test('independent ledgers refuse vacuous inputs, scope shrinkage and altered shared sources',()=>{
 const a=twoSets();expect(()=>reconcileOutcomeMappingSets(a.inventory,[])).toThrow('ledger');
 a.sets[0]!.expectedScopePaths.push('tests/unit/required.test.ts');expect(()=>reconcileOutcomeMappingSets(a.inventory,a.sets)).toThrow('scope');
 const b=twoSets();b.sets.forEach((s,i)=>{s.sources['shared.ts']=String(i);s.ledger.sourceSha256['shared.ts']=createHash('sha256').update(String(i)).digest('hex');});expect(()=>reconcileOutcomeMappingSets(b.inventory,b.sets)).toThrow('source');
});

test('equal shared source pins preserve each independently validated mapping',()=>{
 const {inventory,sets}=twoSets();
 sets.forEach(s=>{s.sources['shared.ts']='shared';s.ledger.sourceSha256['shared.ts']=createHash('sha256').update('shared').digest('hex');});
 const before=JSON.stringify(sets),first=reconcileOutcomeMappings(inventory,sets[0]!.ledger,sets[0]!.features,sets[0]!.sources),report=reconcileOutcomeMappingSets(inventory,sets);
 expect(report.sourceSha256['shared.ts']).toBe(sets[0]!.ledger.sourceSha256['shared.ts']!);
 expect(report.mappings.filter(m=>m.ledger==='first').map(({ledger,...row})=>row)).toEqual(first.mappings);
 expect(JSON.stringify(sets)).toBe(before);
});
test('cross-ledger case keys cannot bypass local scenario ownership',()=>{
 const {inventory,sets}=twoSets();sets[0]!.ledger.mappings[0]!.caseKeys=[sets[1]!.features[0]!.scenarios[0]!.cases[0]!.identityKey];
 expect(()=>reconcileOutcomeMappingSets(inventory,sets)).toThrow('case');
});

test('outcome reconciliation refuses unresolved registrations even outside mapped scopes',async()=>{
 const {mkdtemp,rm,mkdir,writeFile}=await import('node:fs/promises'),{tmpdir}=await import('node:os'),{join}=await import('node:path'),{inventoryNativeTests}=await import('../../scripts/test-inventory.ts');
 const root=await mkdtemp(join(tmpdir(),'bun-unresolved-outcomes-'));
 try{
  await mkdir(join(root,'tests/unit'),{recursive:true});
  await writeFile(join(root,'tests/unit/hidden.test.ts'),`import {describe} from 'bun:test'; const register=()=>{}; describe('hidden',register);`);
  const native=await inventoryNativeTests(root);expect(native.unresolved.length).toBeGreaterThan(0);expect(native.cases).toHaveLength(0);
  const a=sample(),inventory={cases:a.inventory.cases,unresolved:native.unresolved};
  expect(()=>reconcileOutcomeMappings(inventory,a.ledger,a.features,a.sources)).toThrow('Unresolved');
  expect(()=>reconcileOutcomeMappingSets(inventory,[{name:'valid',expectedScopePaths:a.ledger.scopePaths,ledger:a.ledger,features:a.features,sources:a.sources}])).toThrow('Unresolved');
 }finally{await rm(root,{recursive:true,force:true});}
});
test('standalone mapping validation refuses unresolved declaration records',()=>{
 const a=sample();a.inventory.cases[0]!.unresolved.push('unsupported registration');
 expect(()=>reconcileOutcomeMappings(a.inventory,a.ledger,a.features,a.sources)).toThrow('Unresolved');
});

test('XML ledger retains partial links for different fixtures and split edit predicates', async () => {
 const {outcomeMappingReport}=await import('../../scripts/outcome-mappings.ts');
 const report=await outcomeMappingReport(),xml=report.mappings.filter(m=>m.ledger==='xml-values');
 expect(xml).toHaveLength(13);
 const lookup=xml.find(m=>m.testId.endsWith('expanded attribute lookup honours aliases, local rebinding and unqualified attributes'))!;
 expect(lookup.scenarioIds).toEqual(['@id-xml-expanded-attribute-lookup','@id-xml-immutable-namespace-metadata']);
 expect(lookup.caseKeys).toEqual([]);expect(lookup.gaps.join(' ')).toContain('not the dedicated canonical metadata input');
 const edits=xml.filter(m=>m.scenarioIds.includes('@id-xml-apply-edits'));
 expect(edits).toHaveLength(2);expect(edits.every(m=>m.caseKeys.length===0&&m.status==='partial')).toBe(true);
 const implicit=xml.find(m=>m.scenarioIds.includes('@id-xml-implicit-xml-prefix'))!;
 expect(implicit.assertions).toContain("expect(doc.root.attributeNamespaces['xml:lang']).toBe(XML_NS)");
 expect(implicit.assertions).toContain("expect(attribute(doc.root, 'lang', XML_NS)).toBe('en')");
 expect(implicit.assertions.some(a=>a.startsWith('expect(XML_NS)'))).toBe(false);
 expect(xml.every(m=>m.executionCredit===false&&m.gaps.length>0)).toBe(true);
});

test('original PPTX mappings retain aggregate acceptance and bounded custody separately from newer notes records', async () => {
 const {outcomeMappingReport}=await import('../../scripts/outcome-mappings.ts');
 const report=await outcomeMappingReport(),rows=report.mappings.filter(m=>m.ledger==='pptx-core').slice(0,6);
 expect(rows).toHaveLength(6);
 expect(rows[0]!.scenarioIds).toHaveLength(4);expect(rows[0]!.assertions).toHaveLength(3);
 expect(rows[0]!.gaps.join(' ')).toContain('aggregate status and counts');
 const cross=rows.find(m=>m.testId.endsWith('replaces exact anchored text across runs and preserves untouched members'))!;
 expect(cross.gaps.join(' ')).toContain('Only two unrelated saved members');
 const noop=rows.find(m=>m.scenarioIds.includes('@id-pptx-bun-open-save-noop'))!;
 expect(noop.caseKeys).toEqual(['@id-pptx-bun-open-save-noop']);
 expect(noop.gaps.join(' ')).toContain('byte-open disk save');
 expect(rows.every(m=>m.executionCredit===false&&m.status==='partial')).toBe(true);
 expect(rows.flatMap(m=>m.scenarioIds).some(id=>id.includes('go-notes')||id.includes('slide-order'))).toBe(false);
});
