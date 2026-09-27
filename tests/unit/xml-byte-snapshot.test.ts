import {test,expect} from 'bun:test';
import * as api from '../../src/index.ts';
const {XmlByteSnapshot,XmlSnapshot}=api;
const utf8=(s:string)=>new TextEncoder().encode(s);
function encoded(s:string,kind:'utf8'|'le'|'be',bom=true){if(kind==='utf8')return bom?new Uint8Array([239,187,191,...utf8(s)]):utf8(s);const b=new Uint8Array((bom?2:0)+2*s.length);if(bom)b.set(kind==='be'?[254,255]:[255,254]);const v=new DataView(b.buffer);for(let i=0;i<s.length;i++)v.setUint16((bom?2:0)+2*i,s.charCodeAt(i),kind==='le');return b;}

test('byte snapshot owns caller input and empty-removal output with frozen handles',()=>{
 const source=utf8('<r><leaf id="a">one</leaf><leaf id="b">two</leaf></r>'),before=new Uint8Array(source),snapshot=XmlByteSnapshot.parse(source);expect(source).toEqual(before);expect(snapshot.remove([])).toEqual(before);expect(snapshot.elements.map(n=>n.localName)).toEqual(['r','leaf','leaf']);expect(Object.isFrozen(snapshot)).toBe(true);expect(Object.isFrozen(snapshot.elements)).toBe(true);expect(snapshot.elements.every(Object.isFrozen)).toBe(true);
 source.fill(32);const copy=snapshot.remove([]);copy.fill(0);expect(snapshot.remove([])).toEqual(before);expect(snapshot.remove([snapshot.elements[1]!])).toEqual(utf8('<r><leaf id="b">two</leaf></r>'));expect(snapshot.remove([])).toEqual(before);
});

test('UTF8 BOM and UTF16 endian encodings retain exact untouched source and declarations after removal',()=>{
 for(const kind of ['utf8','le','be'] as const)for(const bom of kind==='utf8'?[false,true]:[true]){const declaration=kind==='utf8'?'UTF-8':kind==='le'?'UTF-16LE':'UTF-16BE',source=`<?xml version="1.0" encoding="${declaration}"?>\r\n<r xmlns:q='urn:q'>雪😀\r\n<q:drop a='&amp;'/>\t<keep> e&#x301; </keep><!--tail--></r>`,input=encoded(source,kind,bom),before=new Uint8Array(input),snapshot=XmlByteSnapshot.parse(input),drop=snapshot.elements.find(n=>n.localName==='drop')!;expect(snapshot.remove([])).toEqual(input);expect(snapshot.remove([drop])).toEqual(encoded(source.replace("<q:drop a='&amp;'/>",''),kind,bom));expect(input).toEqual(before);expect(source.slice(drop.start,drop.end)).toBe("<q:drop a='&amp;'/>");}
});

test('generic UTF16 declarations and no declarations preserve BOM endian identity',()=>{
 for(const kind of ['le','be'] as const)for(const declaration of ['', '<?xml version="1.0" encoding="UTF-16"?>']){const text=declaration+'<r><drop/><keep/></r>',bytes=encoded(text,kind),snapshot=XmlByteSnapshot.parse(bytes);expect(snapshot.remove([snapshot.elements[1]!])).toEqual(encoded(declaration+'<r><keep/></r>',kind));}
});

test('removal failures reject root, overlap, duplicates and foreign or forged string/byte handles atomically',()=>{
 const input=utf8('<r><a><b/></a><c/></r>'),before=new Uint8Array(input),s=XmlByteSnapshot.parse(input),other=XmlByteSnapshot.parse(input),string=XmlSnapshot.parse(new TextDecoder().decode(input));for(const targets of [[s.elements[0]!],[s.elements[1]!,s.elements[2]!],[s.elements[3]!,s.elements[3]!],[other.elements[3]!],[string.elements[3]!],[{...s.elements[3]!}]]){expect(()=>s.remove(targets)).toThrow();expect(s.remove([])).toEqual(before);expect(input).toEqual(before);}
 const unsafe=XmlByteSnapshot.parse(utf8('<r>]]<a/>></r>'));expect(()=>unsafe.remove([unsafe.elements[1]!])).toThrow(expect.objectContaining({code:'XML_REMOVAL_UNSAFE'}));expect(unsafe.remove([])).toEqual(utf8('<r>]]<a/>></r>'));
});

test('invalid byte types, malformed encodings, unsupported declarations and BOM mismatches refuse without mutation',()=>{
 for(const value of [null,'<r/>',new ArrayBuffer(4),new DataView(new ArrayBuffer(4)),[60,114,47,62]])expect(()=>XmlByteSnapshot.parse(value as unknown as Uint8Array)).toThrow();
 const cases=[new Uint8Array([0xff]),new Uint8Array([0xc0,0xaf]),new Uint8Array([0xff,0xfe,60]),encoded('<r>\ud800</r>','le'),encoded('<r>\udfff</r>','be'),utf8('<?xml version="1.0" encoding="ISO-8859-1"?><r/>'),utf8('<?xml version="1.0" encoding="UTF-16"?><r/>'),encoded('<?xml version="1.0" encoding="UTF-8"?><r/>','le'),encoded('<?xml version="1.0" encoding="UTF-16LE"?><r/>','be'),encoded('<r/>','le',false),utf8('<!DOCTYPE r><r/>'),utf8('<r>')];
 for(const input of cases){const before=new Uint8Array(input);expect(()=>XmlByteSnapshot.parse(input)).toThrow();expect(input).toEqual(before);}
});

test('byte subviews exclude adjacent data and ignore caller iterator/slice overrides',()=>{
 const source=utf8('XX<r><drop/><keep/></r>YY'),view=source.subarray(2,source.length-2),before=new Uint8Array(view);Object.defineProperty(view,'slice',{value:()=>{throw Error('slice called');}});Object.defineProperty(view,Symbol.iterator,{value:()=>{throw Error('iterator called');}});const s=XmlByteSnapshot.parse(view);expect(s.remove([])).toEqual(before);expect(s.remove([s.elements[1]!])).toEqual(utf8('<r><keep/></r>'));expect(new Uint8Array(view)).toEqual(before);
});

test('byte snapshots enforce byte and decoded XML limits while permitting multibyte non-ASCII input',()=>{
 expect(()=>XmlByteSnapshot.parse(new Uint8Array(32*1024*1024+1))).toThrow(expect.objectContaining({code:'XML_BYTE_LIMIT'}));expect(()=>XmlByteSnapshot.parse(utf8('<r>'+ 'x'.repeat(8*1024*1024)+'</r>'))).toThrow(expect.objectContaining({code:'XML_INPUT_TOO_LARGE'}));const input=utf8('<r>雪😀<a/></r>'),s=XmlByteSnapshot.parse(input);expect(s.remove([s.elements[1]!])).toEqual(utf8('<r>雪😀</r>'));
});

test('shared buffers, detached buffers and duplicate BOMs cannot bypass admission or encoding declarations',()=>{
 const shared=new Uint8Array(new SharedArrayBuffer(8));shared.set(utf8('<r/>'));expect(()=>XmlByteSnapshot.parse(shared)).toThrow(expect.objectContaining({code:'XML_BYTE_SOURCE'}));const bytes=utf8('<r/>');structuredClone(bytes,{transfer:[bytes.buffer]});expect(()=>XmlByteSnapshot.parse(bytes)).toThrow(expect.objectContaining({code:'XML_BYTE_SOURCE'}));
 for(const kind of ['utf8','le','be'] as const){const value=encoded('\ufeff<?xml version="1.0" encoding="ISO-8859-1"?><r/>',kind);expect(()=>XmlByteSnapshot.parse(value)).toThrow(expect.objectContaining({code:'XML_BYTE_ENCODING'}));}
});

test('repeated edits and mutated changed output do not affect source ownership or target reuse',()=>{
 const source=utf8('<r><a/><b/></r>'),snapshot=XmlByteSnapshot.parse(source),result=snapshot.remove([snapshot.elements[1]!]);result.fill(0);expect(snapshot.remove([snapshot.elements[1]!])).toEqual(utf8('<r><b/></r>'));expect(snapshot.remove([snapshot.elements[2]!])).toEqual(utf8('<r><a/></r>'));expect(snapshot.remove([])).toEqual(source);
 const selected=[snapshot.elements[1]!];selected[Symbol.iterator]=function*(){yield snapshot.elements[0]!;return undefined;};expect(snapshot.remove(selected)).toEqual(utf8('<r><b/></r>'));
});

test('canonical byte seed checks input and no-op output independently with corruption controls',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/steps.ts'),{scenarioIds}=await import('../acceptance/xml-byte-snapshot.ts'),{join}=await import('node:path');const path='workflows/xml/editing.feature',source=await Bun.file(join(fixturesRoot(),path)).text(),count=(n:number)=>({implemented:n,planned:0,total:n});const inv={root:'.',features:[selectSharedScenarios(path,source,scenarioIds)],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(4)}};
 const good=await executeAcceptance(inv,bindings,'xml-byte-seed');expect(good.failures).toEqual([]);expect(good.counts.cases.passed).toBe(1);
 for(const field of ['input','output']){const badBindings=bindings.map(b=>b.pattern.test('the XML editor parses a caller-owned byte slice and performs an empty edit')?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);(c.state as {input:Uint8Array;output:Uint8Array})[field as 'input'|'output'].fill(0);}}:b);const bad=await executeAcceptance(inv,badBindings,'xml-byte-corrupt');expect(bad.counts.cases.failed).toBe(1);expect(bad.counts.steps.failed).toBe(1);expect(bad.counts.steps.undefined).toBe(0);expect(bad.counts.steps.ambiguous).toBe(0);}
});

test('intrinsic view bounds reject spoofed length/shared storage without calling caller getters',()=>{
 let called=false;const large=new Uint8Array(32*1024*1024+1);Object.defineProperty(large,'byteLength',{get(){called=true;return 4;}});expect(()=>XmlByteSnapshot.parse(large)).toThrow(expect.objectContaining({code:'XML_BYTE_LIMIT'}));expect(called).toBe(false);
 const shared=new Uint8Array(new SharedArrayBuffer(4));shared.set(utf8('<r/>'));Object.defineProperty(shared,'buffer',{get(){called=true;return new ArrayBuffer(4);}});expect(()=>XmlByteSnapshot.parse(shared)).toThrow(expect.objectContaining({code:'XML_BYTE_SOURCE'}));expect(called).toBe(false);
 const buffer=Buffer.from('<r><drop/></r>'),s=XmlByteSnapshot.parse(buffer);buffer.fill(0);expect(s.remove([])).toEqual(utf8('<r><drop/></r>'));expect(s.remove([s.elements[1]!])).toEqual(utf8('<r></r>'));
});
