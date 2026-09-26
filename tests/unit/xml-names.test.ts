import {describe,expect,test} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {parseXml,attribute} from '../../src/xml/index.ts';
import {parseFeature,executeAcceptance,newAcceptanceRunId,type AcceptanceInventory} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/xml-names.ts';
describe('XML QName component admission',()=>{
 test('rejects numeric, punctuation and combining-mark leading local names and declared prefixes',()=>{
  for(const xml of ['<p:1 xmlns:p="urn:p"/>','<p:-x xmlns:p="urn:p"/>','<p:.x xmlns:p="urn:p"/>','<p:\u0301x xmlns:p="urn:p"/>','<r xmlns:p="urn:p" p:1="x"/>','<r xmlns:1="urn:p"/>','<r xmlns:-p="urn:p"/>','<r xmlns:p:q="urn:p"/>'])expect(()=>parseXml(xml)).toThrow();
 });
 test('retains legal Unicode NCNames and exact source offsets',()=>{
  for(const name of ['π','名','é','𐐀','a1','_x','a-b','a.b','a\u0301']){const xml=`<π:${name} xmlns:π="urn:n" π:${name}="yes"/>`;const d=parseXml(xml);expect(d.root.localName).toBe(name);expect(attribute(d.root,name,'urn:n')).toBe('yes');expect(xml.slice(d.root.start,d.root.end)).toBe(xml);}
 });
 test('outside-root XML whitespace is lexical and NBSP is never whitespace',()=>{
  for(const source of ['\u00a0<r/>','<r/>\u00a0','&#32;<r/>','<r/>&#x20;'])expect(()=>parseXml(source)).toThrow();
  expect(parseXml(' \r\n\t<r/> \r\n\t').root.localName).toBe('r');
 });
 test('literal whitespace normalises before character reference decoding',()=>{
  const d=parseXml('<r a="x\r\ny\tz&#9;&#10;&#13;">a\r\nb\rc<![CDATA[d\r\ne\rf]]></r>');
  expect(d.root.attributes.a).toBe('x y z\t\n\r');expect(d.root.text).toBe('a\nb\ncd\ne\nf');
 });
 test('executes the six QName/whitespace acceptance cases',async()=>{
  const path='workflows/xml/names.feature',text=await Bun.file(join(fixturesRoot(),path)).text();
  const f=parseFeature('references/fixtures-ooxml/'+path,text.replace(/^@planned/m,'@implemented @bun'));const cases=f.scenarios.flatMap(s=>s.cases);const count=(n:number)=>({implemented:n,planned:0,total:n});
  const inventory:AcceptanceInventory={root:'.',features:[f],counts:{features:count(1),scenarios:count(f.scenarios.length),cases:count(cases.length),steps:count(cases.reduce((n,c)=>n+c.steps.length,0))}};
  const result=await executeAcceptance(inventory,bindings,newAcceptanceRunId());expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(6);
 });
});
