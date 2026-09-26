import {test,expect} from 'bun:test';
import {xmlEquivalent} from '../../src/index.ts';
import {utf8,utf16} from '../fixtures/admission.ts';
const eq=(a:string,b:string)=>xmlEquivalent(utf8(a),utf8(b));
test('XML comparison uses expanded names and decoded values, not prefix spelling',()=>{
 expect(eq('<a xmlns="u" z="2" x="&amp;"/>','<p:a xmlns:p="u" x="&#38;" z="2"></p:a>')).toBe(true);
 expect(eq('<a xmlns:p="u" p:x="yes"/>','<a xmlns:q="u" q:x="yes"/>')).toBe(true);
 expect(eq('<a x="yes"/>','<a xmlns:p="u" p:x="yes"/>')).toBe(false);
});
test('comparison preserves text, tails, child order and markup content',()=>{
 for(const [a,b] of [['<a><b/> tail</a>','<a> tail<b/></a>'],['<a><b/><c/></a>','<a><c/><b/></a>'],['<a><!--old--></a>','<a><!--new--></a>'],['<a><?one x?></a>','<a><?two x?></a>'],['<?one x?><a/>','<?one y?><a/>'],['<a/><!--one-->','<a/><!--two-->'],['<a> x </a>','<a>x</a>']])expect(eq(a!,b!)).toBe(false);
 expect(eq('<a>x<![CDATA[&y]]>z</a>','<a>x&amp;yz</a>')).toBe(true);
 expect(eq('<a>one\r\ntwo</a>','<a>one\ntwo</a>')).toBe(true);
 expect(eq('<a>&#13;</a>','<a>\r</a>')).toBe(false);
});
test('OPC relationship order is ignored only for unambiguous relation collections',()=>{
 const ns='http://schemas.openxmlformats.org/package/2006/relationships',a='<Relationship Id="a" Target="a.xml"/>',b='<Relationship Id="b" Target="b.xml"/>';
 expect(eq(`<Relationships xmlns="${ns}">${a}${b}</Relationships>`,`<Relationships xmlns="${ns}">${b}${a}</Relationships>`)).toBe(true);
 expect(eq(`<Relationships xmlns="other">${a}${b}</Relationships>`,`<Relationships xmlns="other">${b}${a}</Relationships>`)).toBe(false);
 expect(eq(`<Relationships xmlns="${ns}">${a}${a}</Relationships>`,`<Relationships xmlns="${ns}">${a}${a}</Relationships>`)).toBe(false);
 expect(eq(`<Relationships xmlns="${ns}">${a}text${b}</Relationships>`,`<Relationships xmlns="${ns}">${b}text${a}</Relationships>`)).toBe(false);
});
test('prefix-valued attributes retain bindings including inherited and implicit namespaces',()=>{
 expect(eq('<a xmlns:p="u" value="p:x"/>','<a xmlns:p="v" value="p:x"/>')).toBe(false);
 expect(eq('<a xmlns:p="u"><b value="p:x"/></a>','<a xmlns:p="v"><b value="p:x"/></a>')).toBe(false);
 expect(eq('<a xmlns:p="u" value="p:x"/>','<a xmlns:q="u" value="q:x"/>')).toBe(false);
 expect(eq('<a value="xml:lang"/>','<a value="xml:lang"/>')).toBe(true);
});
test('malformed, DTD, encoding conflicts and missing QName bindings never compare true',()=>{
 for(const x of ['broken','<a>','<!DOCTYPE a><a/>','<!DOCTYPE a [<!ENTITY e "x">]><a>&e;</a>','<a xmlns:p="u" p:x="1" p:x="2"/>','<a value="missing:x"/>'])expect(eq(x,x)).toBe(false);
 expect(xmlEquivalent(Uint8Array.of(255),Uint8Array.of(255))).toBe(false);
 expect(eq('<?xml version="1.0" encoding="UTF-16"?><a/>','<a/>')).toBe(false);
});
test('comparison supports matching UTF encodings without changing either input',()=>{
 const a=utf16('<?xml version="1.0" encoding="UTF-16"?><a>雪</a>'),b=utf8('<a>雪</a>'),beforeA=a.slice(),beforeB=b.slice();expect(xmlEquivalent(a,b)).toBe(true);expect(a).toEqual(beforeA);expect(b).toEqual(beforeB);
});

test('comparison retains standalone intent and refuses unsupported XML versions',()=>{
 expect(eq('<?xml version="1.0"?><a/>','<a/>')).toBe(true);
 expect(eq('<?xml version="1.0" standalone="yes"?><a/>','<a/>')).toBe(false);
 expect(eq('<?xml version="1.1"?><a/>','<?xml version="1.1"?><a/>')).toBe(false);
 expect(eq('<?xml version="1.0" standalone="yes"?><a/>','<?xml-standalone yes?><a/>')).toBe(false);
});
test('MC prefix lists and nested rebinding are not lost during namespace normalization',()=>{
 const mc='http://schemas.openxmlformats.org/markup-compatibility/2006';
 expect(eq(`<a xmlns:mc="${mc}" xmlns:p="u" mc:Ignorable="p"/>`,`<a xmlns:mc="${mc}" xmlns:p="v" mc:Ignorable="p"/>`)).toBe(false);
 expect(eq('<a xmlns:p="u"><b xmlns:p="v" value="p:x"/></a>','<a xmlns:p="u"><b value="p:x"/></a>')).toBe(false);
});
test('all ten shared comparator cases execute with exact booleans and input custody',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{join}=await import('node:path'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/xml-comparison.ts');
 const path='workflows/xml/comparison.feature',ids=['@id-xml-comparison-prefix-and-opc-order','@id-xml-comparison-significant-content','@id-xml-comparison-prefix-attribute-binding','@id-xml-comparison-unsafe-input','@id-xml-comparison-processing-instructions-and-comments'];
 const feature=selectSharedScenarios(path,await Bun.file(join(fixturesRoot(),path)).text(),ids),count=(n:number)=>({implemented:n,planned:0,total:n});
 const result=await executeAcceptance({root:'.',features:[feature],counts:{features:count(1),scenarios:count(5),cases:count(10),steps:count(40)}},bindings,'comparison-unit');expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(10);
});
test('deep segmented comparison is bounded without repeated ancestor text and depth violations return false',()=>{
 const source='<a>'.repeat(256)+('x'.repeat(4096)+'<!--segment-->').repeat(64)+'</a>'.repeat(256);expect(eq(source,source)).toBe(true);
 const deep='<a>'.repeat(257)+'</a>'.repeat(257);expect(eq(deep,deep)).toBe(false);
});

test('unprefixed xsi:type values retain the default namespace that resolves their QName',()=>{
 const xsi='http://www.w3.org/2001/XMLSchema-instance';
 expect(eq(`<e:r xmlns:e="urn:element" xmlns="urn:one" xmlns:xsi="${xsi}" xsi:type="Type"/>`,`<e:r xmlns:e="urn:element" xmlns="urn:two" xmlns:xsi="${xsi}" xsi:type="Type"/>`)).toBe(false);
});

test('MC Choice Requires prefix lists retain bindings even without colon-valued attributes',()=>{
 const mc='http://schemas.openxmlformats.org/markup-compatibility/2006';
 expect(eq(`<mc:Choice xmlns:mc="${mc}" xmlns:p="urn:one" Requires="p"/>`,`<mc:Choice xmlns:mc="${mc}" xmlns:p="urn:two" Requires="p"/>`)).toBe(false);
});

test('MC MustUnderstand lists and QName-like text retain in-scope namespace bindings',()=>{
 const mc='http://schemas.openxmlformats.org/markup-compatibility/2006';
 expect(eq(`<a xmlns:mc="${mc}" xmlns:p="urn:one" mc:MustUnderstand="p"/>`,`<a xmlns:mc="${mc}" xmlns:p="urn:two" mc:MustUnderstand="p"/>`)).toBe(false);
 expect(eq(`<a xmlns:mc="${mc}" mc:MustUnderstand="missing"/>`,`<a xmlns:mc="${mc}" mc:MustUnderstand="missing"/>`)).toBe(false);
 expect(eq('<a xmlns:p="urn:one">p:x</a>','<a xmlns:p="urn:two">p:x</a>')).toBe(false);
 expect(eq('<a xmlns:p="urn:one"><b>p<![CDATA[:x]]></b></a>','<a xmlns:p="urn:two"><b>p:x</b></a>')).toBe(false);
 expect(eq('<a xmlns:p="urn:one">p<![CDATA[:x]]></a>','<a xmlns:p="urn:one">p:x</a>')).toBe(true);
});

test('processing instruction spacing is retained rather than trimmed by parser events',()=>{
 expect(eq('<?p x?><a/>','<?p  x?><a/>')).toBe(false);
 expect(eq('<a><?p\tx?></a>','<a><?p x?></a>')).toBe(false);
 expect(eq('<?p?><a/>','<?p ?><a/>')).toBe(false);
 expect(eq('<?p\r\nx?><a/>','<?p\nx?><a/>')).toBe(true);
});
