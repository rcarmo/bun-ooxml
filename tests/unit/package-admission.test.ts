import {test,expect} from 'bun:test';
import {admitPackage,OoxmlError,OpcPackage} from '../../src/index.ts';
import {readZip} from '../../src/opc/zip.ts';
import {admissionZip,utf8,utf16,BZIP2_A} from '../fixtures/admission.ts';
const xml='<a/>';
test('admission accepts XML without an OPC graph and returns detached payloads',async()=>{
 const bytes=admissionZip([['a.xml',utf8(xml)],['opaque.bin',Uint8Array.of(0xff,0x00)]]),before=bytes.slice();
 const parts=admitPackage(bytes);expect(parts.get('a.xml')).toEqual(utf8(xml));expect(parts.get('opaque.bin')).toEqual(Uint8Array.of(0xff,0x00));expect(bytes).toEqual(before);
 parts.get('a.xml')!.fill(0);expect(bytes).toEqual(before);expect(admitPackage(bytes).get('a.xml')).toEqual(utf8(xml));
 await expect(OpcPackage.open(bytes)).rejects.toMatchObject({code:'opc-content-types-missing'});
});
test('admission preserves UTF-8 BOM, UTF-16LE and UTF-16BE XML payloads exactly',()=>{
 for(const payload of [Uint8Array.from([0xef,0xbb,0xbf,...utf8(xml)]),utf16('<?xml version="1.0" encoding="UTF-16"?><a>雪</a>'),utf16('<?xml version="1.0" encoding="UTF-16BE"?><a>雪</a>',true)]){
  const bytes=admissionZip([['a.xml',payload]]),before=bytes.slice();expect(admitPackage(bytes).get('a.xml')).toEqual(payload);expect(bytes).toEqual(before);
 }
});
test('admission validates every XML and relationship member but keeps unrelated bytes opaque',()=>{
 for(const name of ['a.xml','A.XML','_rels/.rels','custom.RELS']){
  const bytes=admissionZip([['valid.xml',utf8(xml)],[name,utf8('<broken>')]]),before=bytes.slice();expect(()=>admitPackage(bytes)).toThrow(OoxmlError);expect(bytes).toEqual(before);
 }
 expect(admitPackage(admissionZip([['opaque.bin',utf8('<broken>')]])).get('opaque.bin')).toEqual(utf8('<broken>'));
});
test('admission refuses DTD and malformed XML with parser-specific reasons without changing bytes',()=>{
 for(const [payload,code]of [[utf8('<!DOCTYPE a [<!ENTITY e "text">]><a>&e;</a>'),'XML_DTD_FORBIDDEN'],[utf16('<!DOCTYPE a><a/>'),'XML_DTD_FORBIDDEN'],[utf8('<broken>'),'XML_MALFORMED']] as const){
  const bytes=admissionZip([['a.xml',payload]]),before=bytes.slice();expect(()=>admitPackage(bytes)).toThrow(expect.objectContaining({code}));expect(bytes).toEqual(before);
 }
});
test('admission rejects invalid and contradictory byte encodings rather than decoding lossily',()=>{
 for(const payload of [Uint8Array.of(0xff),utf16('<a/>').slice(0,-1),utf8('<?xml version="1.0" encoding="UTF-16"?><a/>'),utf16('<?xml version="1.0" encoding="UTF-8"?><a/>'),utf8('<?xml version="1.0" encoding="ISO-8859-1"?><a/>')]){
  const bytes=admissionZip([['a.xml',payload]]),before=bytes.slice();expect(()=>admitPackage(bytes)).toThrow(expect.objectContaining({code:'opc-xml-encoding'}));expect(bytes).toEqual(before);
 }
});
test('admission resource controls have successful exact-boundary controls and typed refusals',()=>{
 const bytes=admissionZip([['a.xml',utf8(xml)]],0),before=bytes.slice();
 expect(admitPackage(bytes,{maxArchiveBytes:bytes.length,maxEntries:1,maxEntryBytes:4,maxTotalBytes:4,maxCompressionRatio:1}).size).toBe(1);
 for(const [limits,code]of [[{maxArchiveBytes:bytes.length-1},'zip-archive-too-large'],[{maxEntries:0},'zip-too-many-entries'],[{maxEntryBytes:3},'zip-entry-too-large'],[{maxTotalBytes:3},'zip-total-too-large'],[{maxCompressionRatio:0.5},'zip-compression-ratio-exceeded']] as const){expect(()=>admitPackage(bytes,limits)).toThrow(expect.objectContaining({code}));expect(bytes).toEqual(before);}
});
test('invalid admission limit configuration refuses rather than disabling bounds',()=>{
 const bytes=admissionZip([['a.xml',utf8(xml)]]);
 for(const limits of [{maxEntries:NaN},{maxEntryBytes:Infinity},{maxTotalBytes:-1},{maxArchiveBytes:1.5},{maxCompressionRatio:0},{maxCompressionRatio:NaN}])expect(()=>admitPackage(bytes,limits)).toThrow(expect.objectContaining({code:'package-admission-limit-invalid'}));
});
test('ordered ZIP fixtures reach typed structural refusals, with valid DEFLATE and BZIP2 method framing',()=>{
 const deflated=admissionZip([['a.xml',utf8('<a>'+' '.repeat(10000)+'</a>')]],8);expect(readZip(deflated).get('a.xml')!.length).toBe(10007);
 const bzip=admissionZip([['a.xml',utf8(xml)]],12);expect([...bzip.slice(35,35+BZIP2_A.length)]).toEqual([...BZIP2_A]);expect(()=>admitPackage(bzip)).toThrow(expect.objectContaining({code:'zip-method-unsupported'}));
 for(const [entries,code] of [
  [[['a.xml','<a/>'],['a.xml','<b/>']],'zip-duplicate-entry'],[[['../a.xml','<a/>']],'zip-name-invalid'],[[['/a.xml','<a/>']],'zip-name-invalid'],[[['x\\a.xml','<a/>']],'zip-name-invalid'],[[['a/','payload']],'zip-directory-entry-invalid'],
 ] as const){const bytes=admissionZip(entries.map(([n,s])=>[n,utf8(s)] as const)),before=bytes.slice();expect(()=>admitPackage(bytes)).toThrow(expect.objectContaining({code}));expect(bytes).toEqual(before);}
});

test('all thirteen shared admission cases execute concrete refusals with positive controls',async()=>{
 const {join}=await import('node:path'),{fixturesRoot}=await import('../../scripts/fixture-inputs.ts');
 const {selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/package-admission.ts');
 const features=[];
 for(const [path,ids]of [['workflows/package/zip-admission.feature',['@id-package-admission-unsafe-members','@id-package-admission-resource-limits','@id-package-admission-unsupported-compression']],['workflows/package/xml-member-admission.feature',['@id-package-admission-unsafe-xml-members']]] as const){
  features.push(selectSharedScenarios(path,await Bun.file(join(fixturesRoot(),path)).text(),[...ids]));
 }
 const count=(n:number)=>({implemented:n,planned:0,total:n});const result=await executeAcceptance({root:'.',features,counts:{features:count(2),scenarios:count(4),cases:count(13),steps:count(40)}},bindings,'unit-admission');
 expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(13);expect(result.counts.steps.passed).toBe(40);
});

test('admission checks CRC, refuses external DTDs and handles archive views without borrowing backing bytes',()=>{
 const valid=admissionZip([['a.xml',utf8('<a/>')],['empty/',new Uint8Array()]]),storage=new Uint8Array(valid.length+16);storage.set(valid,8);
 const view=storage.subarray(8,8+valid.length),before=storage.slice(),parts=admitPackage(view);expect([...parts.keys()]).toEqual(['a.xml']);expect(storage).toEqual(before);parts.get('a.xml')!.fill(0);expect(storage).toEqual(before);
 const corrupt=admissionZip([['a.xml',utf8('<a/>')]]);corrupt[35]=0x20;const snapshot=corrupt.slice();expect(()=>admitPackage(corrupt)).toThrow(expect.objectContaining({code:'zip-crc-mismatch'}));expect(corrupt).toEqual(snapshot);
 for(const payload of [utf8('<!DOCTYPE a SYSTEM "https://example.invalid/no-fetch"><a/>'),utf16('<!DOCTYPE a><a/>',true)])expect(()=>admitPackage(admissionZip([['a.xml',payload]]))).toThrow(expect.objectContaining({code:'XML_DTD_FORBIDDEN'}));
});

test('syntax-only XML validation handles deep segmented text without ancestor text materialisation',async()=>{
 const {validateXml,parseXml}=await import('../../src/xml/index.ts');
 const source='<a>'.repeat(256)+('x'.repeat(8192)+'<!--segment-->').repeat(256)+'</a>'.repeat(256);
 expect(validateXml(source)).toBeUndefined();
 expect(()=>validateXml('<a>'.repeat(257)+'x'+'</a>'.repeat(257))).toThrow(expect.objectContaining({code:'XML_DEPTH_LIMIT'}));
 expect(()=>validateXml('<a>&unknown;</a>')).toThrow(expect.objectContaining({code:'XML_ENTITY_FORBIDDEN'}));
 const bytes=admissionZip([['deep.xml',utf8(source)]],0);expect(admitPackage(bytes).get('deep.xml')).toEqual(utf8(source));
 // Public parser still provides descendant text for editors.
 expect(parseXml('<a>x<b>y</b>z</a>').root.text).toBe('xyz');
});
