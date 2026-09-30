import {test,expect} from 'bun:test';
import {xmlEquivalent} from '../../src/xml/comparison.ts';
const bytes=(s:string)=>new TextEncoder().encode(s);
const ns='http://schemas.openxmlformats.org/package/2006/relationships';
const compare=(a:string,b:string)=>xmlEquivalent(bytes(a),bytes(b));
test('valid OPC relationship Type URIs are literals, not unknown QName prefixes',()=>{
 const a=`<Relationships xmlns="${ns}"><Relationship Id="r1" Type="urn:test/a" Target="a.xml"/><Relationship Id="r2" Type="urn:test/b" Target="b.xml"/></Relationships>`;
 const b=`<p:Relationships xmlns:p="${ns}"><p:Relationship Target="b.xml" Type="urn:test/b" Id="r2"/><p:Relationship Target="a.xml" Type="urn:test/a" Id="r1"/></p:Relationships>`;
 expect(compare(a,b)).toBe(true);
 expect(compare(a,b.replace('urn:test/b','urn:test/c'))).toBe(false);
 expect(compare(a,b.replace('b.xml','c.xml'))).toBe(false);
 expect(compare(a,b.replace('Id="r2"','Id="r1"'))).toBe(false);
 expect(compare('<a value="urn:test/a"/>','<a value="urn:test/a"/>')).toBe(false);
 expect(compare('<a xmlns:p="urn:one" value="p:x"/>','<a xmlns:p="urn:two" value="p:x"/>')).toBe(false);
});
test('relationship External Target URIs also compare literally; mode and Type differences remain significant',()=>{
 const a=`<Relationships xmlns="${ns}"><Relationship Id="r1" Type="https://example.org/relation" Target="https://example.org/item" TargetMode="External"/></Relationships>`;
 const b=a.replace('<Relationships xmlns=', '<p:Relationships xmlns:p=').replace('<Relationship Id=', '<p:Relationship Id=').replace('</Relationships>','</p:Relationships>');
 expect(compare(a,b)).toBe(true);
 expect(compare(a,b.replace('External','Internal'))).toBe(false);
});
