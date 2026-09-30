import {test, expect} from 'bun:test';
import {readFileSync} from 'node:fs';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {selectSharedScenarios, executeAcceptance} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/steps.ts';
import {parseXml, attribute} from '../../src/xml/index.ts';
import {OoxmlError, classifyXmlParseFailure} from '../../src/errors.ts';

const ids=['@id-xml-prototype-safe-attributes','@id-xml-immutable-namespace-metadata','@id-xml-typed-parse-error'];
const root = fixturesRoot(), path='workflows/xml/parsing.feature';
const text=readFileSync(join(root,path),'utf8');
const generalized=text.includes('@profile-xml-model-safety');
async function run(change=(s:string)=>s, active=bindings) {
  const feature=selectSharedScenarios(path,change(text),ids);
  const count=(n:number)=>({implemented:n,planned:0,total:n});
  return executeAcceptance({root:'.',features:[feature],counts:{features:count(1),scenarios:count(3),cases:count(3),steps:count(15)}},active,'xml-model-safety');
}
(generalized?test:test.skip)('candidate generalized XML IDs execute exactly three cases and fifteen steps',async()=>{
  const r=await run();expect(r.failures).toEqual([]);expect(r.counts.cases.passed).toBe(3);
  const executed=r.features.flatMap(f=>f.scenarios.flatMap(s=>s.cases)).filter(c=>c.lifecycle==='implemented');
  expect(executed.reduce((n,c)=>n+c.steps.length,0)).toBe(15);
  expect(executed.every(c=>c.steps.every(s=>s.status==='passed'))).toBe(true);
});
(generalized?test:test.skip)('candidate XML bindings reject wrong literal keys, mutable metadata and unrelated parse failures',async()=>{
  for(const change of [
    (s:string)=>s.replace('__proto__=polluted and constructor=safe','__proto__=wrong and constructor=safe'),
    (s:string)=>s.replace('urn:a and outer','urn:a and wrong'),
  ]){const r=await run(change);expect(r.counts.cases.failed).toBeGreaterThan(0);}
  for(const mutate of [
    (s:any)=>{if(s.document)s.document.root.attributeNamespaces={...s.document.root.attributeNamespaces};},
    (s:any)=>{if(s.error)s.error=new Error('unrelated');},
    (s:any)=>{if(s.document)s.document.root.name='changed';},
  ]){
    const active=bindings.map(b=>b.pattern.test('the XML values input is parsed')?{...b,run:async(c:Record<string,unknown>,...cap:string[])=>{await b.run(c,...cap);mutate(c.state);}}:b);
    const r=await run(s=>s,active);expect(r.counts.cases.failed).toBeGreaterThan(0);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);
  }
});
test('native parser safety controls preserve stronger Bun guarantees and classify only known syntax failures',()=>{
  expect(classifyXmlParseFailure(new Error('other'))).toBeUndefined();
  expect(classifyXmlParseFailure(new OoxmlError('OTHER','other'))).toBeUndefined();
  expect(classifyXmlParseFailure(new OoxmlError('XML_INPUT_TOO_LARGE','resource'))).toBeUndefined();
  expect(parseXml('<a/>').root.name).toBe('a');
  let err:unknown;try{parseXml('<a></b>');}catch(e){err=e;}
  expect(classifyXmlParseFailure(err)).toBe('malformed-xml');
  const attrs=parseXml('<r __proto__="polluted" constructor="safe"/>').root.attributes;
  expect(Object.getPrototypeOf(attrs)).toBeNull();expect(Object.keys(attrs).sort()).toEqual(['__proto__','constructor']);
  expect(Object.hasOwn(attrs,'constructor')).toBe(true);
  const r=parseXml('<r xmlns:a="urn:a" a:id="outer"/>').root,m=r.attributeNamespaces;
  expect(Object.isFrozen(m)).toBe(true);expect(Object.getPrototypeOf(m)).toBeNull();
  expect(Reflect.set(m,'a:id','urn:changed')).toBe(false);expect(attribute(r,'id','urn:a')).toBe('outer');
  expect(Reflect.deleteProperty(m,'a:id')).toBe(false);expect(r.attributeNamespaces['a:id']).toBe('urn:a');
});
