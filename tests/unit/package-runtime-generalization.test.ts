import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {selectSharedScenarios,executeAcceptance} from '../../scripts/gherkin.ts';
import {bindings,cleanup} from '../acceptance/steps.ts';
import {OoxmlError,classifyPackageFailure,OpcPackage} from '../../src/index.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';
const root=fixturesRoot(),opc='workflows/package/preservation.feature',zip='workflows/package/zip32.feature';
const generalized=(await Bun.file(join(root,zip)).text()).includes('@profile-zip32-refusal-reasons');
const opcIds=['@id-bun-opc-open-refusal','@id-bun-opc-save-invalid-target-custody','@id-bun-opc-symlink-destination-refusal'];
const zipIds=['@id-bun-zip32-reader-refusal','@id-bun-zip32-writer-refusal','@id-bun-zip32-configured-bounds'];
async function run(active=bindings){
 const count=(n:number)=>({implemented:n,planned:0,total:n});
 try{return await executeAcceptance({root:'.',features:[selectSharedScenarios(opc,await Bun.file(join(root,opc)).text(),opcIds),selectSharedScenarios(zip,await Bun.file(join(root,zip)).text(),zipIds)],counts:{features:count(2),scenarios:count(6),cases:count(26),steps:count(160)}},active,'package-reasons');}finally{await cleanup();}
}
(generalized?test:test.skip)('candidate six OPC/ZIP IDs execute all26 cases with reason and custody assertions',async()=>{
 const r=await run();expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(26);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);
 const rows=r.features.flatMap(f=>f.scenarios.flatMap(s=>s.cases)).filter(c=>c.lifecycle==='implemented');expect(rows).toHaveLength(26);expect(rows.every(c=>c.steps.every(s=>s.status==='passed'))).toBe(true);
});
(generalized?test:test.skip)('every selected reason distinguishes a wrong structured code and unrelated errors',async()=>{
 for(const replacement of [()=>new OoxmlError('zip-wrong','wrong'),()=>new Error('zip-crc-mismatch'),()=>new OoxmlError('zip-crc-mismatch','wrong row')]){
  const active=bindings.map(b=>b.pattern.test('the package editor opens its archive bytes')||b.pattern.test('the ZIP reader reads the sample with default limits')?{...b,run:async(c:Record<string,unknown>,...cap:string[])=>{await b.run(c,...cap);(c.state as {error:unknown}).error=replacement();}}:b);
  const r=await run(active);expect(r.counts.cases.failed).toBeGreaterThan(0);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);
 }
});
test('native package/ZIP positive siblings succeed and classification never relies on diagnostic text',async()=>{
 const encode=(s:string)=>new TextEncoder().encode(s);
 const entries=new Map([['a.bin',encode('A')],['b.bin',encode('bb')]]),bytes=writeZip(entries),got=readZip(bytes);
 expect([...got.keys()]).toEqual(['a.bin','b.bin']);expect([...got.values()]).toEqual([...entries.values()]);
 expect(classifyPackageFailure(new Error('zip-crc-mismatch'))).toBeUndefined();expect(classifyPackageFailure(new OoxmlError('OTHER','zip-crc-mismatch'))).toBeUndefined();
 expect(classifyPackageFailure(new OoxmlError('zip-crc-mismatch','unrelated wording'))).toBe('zip-crc-mismatch');
 const envelope=writeZip(new Map([
 ['[Content_Types].xml',encode('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')],
 ['_rels/.rels',encode('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')],
 ['word/document.xml',encode('<document>Alpha</document>')],
 ]));const p=await OpcPackage.open(envelope);expect(p.text('word/document.xml')).toBe('<document>Alpha</document>');expect(p.toBytes()).toEqual(envelope);
});
