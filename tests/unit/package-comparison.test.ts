import {test,expect} from 'bun:test';
import {comparePackageArchives,OpcPackage,Document} from '../../src/index.ts';
import {diffPackages} from '../../src/opc/diff.ts';
import {admissionZip,utf8,utf16} from '../fixtures/admission.ts';
import {readZip,writeZip} from '../../src/opc/zip.ts';
const archive=(entries:ReadonlyArray<readonly[string,string]>)=>admissionZip(entries.map(([name,text])=>[name,utf8(text)]));
test('archive comparison separates equivalent XML, changed binary, additions and removals without requiring an OPC graph',()=>{
 const before=archive([['a.xml','<a xmlns="u"><b x="1"/></a>'],['b.bin','x'],['removed.bin','gone']]);
 const after=archive([['a.xml','<p:a xmlns:p="u"><p:b x="1"/></p:a>'],['b.bin','y'],['c.bin','z']]);
 const b=before.slice(),a=after.slice();
 expect(comparePackageArchives(before,after)).toEqual({added:['c.bin'],removed:['removed.bin'],changed:['b.bin'],equivalent_xml:['a.xml'],unchanged:[]});expect(before).toEqual(b);expect(after).toEqual(a);
});
test('exact payload equality is unchanged even when ZIP storage or metadata differ',()=>{
 const entries=[['a.xml',utf8('<a/>')],['opaque.bin',utf8('exact')]] as const,before=admissionZip(entries,0),after=admissionZip([...entries].reverse(),8);
 expect(before).not.toEqual(after);expect(comparePackageArchives(before,after)).toEqual({added:[],removed:[],changed:[],equivalent_xml:[],unchanged:['a.xml','opaque.bin']});
});
test('only XML suffixes receive equivalence checks and semantically significant XML stays changed',()=>{
 const a='<a xmlns="u"><b x="1"/></a>',b='<p:a xmlns:p="u"><p:b x="1"/></p:a>';
 const before=archive([['same.bin',a],['ITEM.XML',a],['_rels/.rels',a],['value.xml','<a> x </a>'],['comment.xml','<!--old--><a/>']]);
 const after=archive([['same.bin',b],['ITEM.XML',b],['_rels/.rels',b],['value.xml','<a>x</a>'],['comment.xml','<!--new--><a/>']]);
 expect(comparePackageArchives(before,after)).toEqual({added:[],removed:[],changed:['comment.xml','same.bin','value.xml'],equivalent_xml:['ITEM.XML','_rels/.rels'],unchanged:[]});
});
test('input admission refuses unsafe XML anywhere even if identical, added or removed',()=>{
 const valid=archive([['a.xml','<a/>']]),invalid=archive([['bad.xml','<!DOCTYPE a><a/>']]);
 for(const [before,after]of [[invalid,invalid],[valid,invalid],[invalid,valid]]){const b=before!.slice(),a=after!.slice();expect(()=>comparePackageArchives(before!,after!)).toThrow(expect.objectContaining({code:'XML_DTD_FORBIDDEN'}));expect(before).toEqual(b);expect(after).toEqual(a);}
});
test('each archive passes ZIP integrity and independent resource budgets before classification',()=>{
 const valid=archive([['a.xml','<a/>']]),large=archive([['a.xml','<a/>'],['b.bin','xx']]),corrupt=valid.slice();corrupt[35]=0x20;
 expect(()=>comparePackageArchives(valid,corrupt)).toThrow(expect.objectContaining({code:'zip-crc-mismatch'}));
 expect(()=>comparePackageArchives(valid,large,{maxEntries:1})).toThrow(expect.objectContaining({code:'zip-too-many-entries'}));
 expect(()=>comparePackageArchives(large,valid,{maxEntries:1})).toThrow(expect.objectContaining({code:'zip-too-many-entries'}));
 expect(comparePackageArchives(valid,valid,{maxEntries:1,maxEntryBytes:4,maxTotalBytes:4}).unchanged).toEqual(['a.xml']);
 expect(()=>comparePackageArchives(valid,valid,{maxEntries:NaN})).toThrow(expect.objectContaining({code:'package-admission-limit-invalid'}));
});
test('UTF encoding differences can be equivalent but namespace-sensitive values cannot',()=>{
 const before=admissionZip([['text.xml',utf16('<?xml version="1.0" encoding="UTF-16"?><a>雪</a>')],['q.xml',utf8('<a xmlns:p="one" v="p:x"/>')]]);
 const after=archive([['text.xml','<a>雪</a>'],['q.xml','<a xmlns:p="two" v="p:x"/>']]);expect(comparePackageArchives(before,after)).toEqual({added:[],removed:[],changed:['q.xml'],equivalent_xml:['text.xml'],unchanged:[]});
});
test('category lists are sorted, disjoint, exhaustive and detached across comparisons',()=>{
 const before=archive([['z.bin','same'],['r.bin','gone'],['a.xml','<a/>'],['m.bin','old']]),after=archive([['b.bin','new'],['m.bin','new'],['a.xml','<a></a>'],['z.bin','same']]);
 const result=comparePackageArchives(before,after),all=Object.values(result).flat();expect(new Set(all).size).toBe(all.length);expect([...all].sort()).toEqual(['a.xml','b.bin','m.bin','r.bin','z.bin']);
 const reverse=comparePackageArchives(after,before);expect(reverse.added).toEqual(result.removed);expect(reverse.removed).toEqual(result.added);expect(reverse.changed).toEqual(result.changed);expect(reverse.equivalent_xml).toEqual(result.equivalent_xml);
 result.changed.push('fake');expect(comparePackageArchives(before,after).changed).toEqual(['m.bin']);
});
test('empty archives and empty directory members do not invent file differences',()=>{
 const empty=admissionZip([]),directory=admissionZip([['empty/',new Uint8Array()]]);expect(comparePackageArchives(empty,directory)).toEqual({added:[],removed:[],changed:[],equivalent_xml:[],unchanged:[]});
});
test('existing OPC byte/content-type diff retains its behavior for XML whitespace changes',async()=>{
 const d=Document.create();d.addParagraph('Keep');const original=d.package.toBytes(),parts=readZip(original),source=new TextDecoder().decode(parts.get('word/document.xml')!);parts.set('word/document.xml',utf8(source.replace('</w:document>','\n</w:document>')));const changed=writeZip(parts);
 const before=await OpcPackage.open(original),after=await OpcPackage.open(changed);expect(diffPackages(before,after).changed).toContain('word/document.xml');
 // Whitespace inside the document element remains significant in this conservative profile.
 expect(comparePackageArchives(original,changed).changed).toContain('word/document.xml');
});

test('the exact shared package semantic-diff case compares concrete table inputs',async()=>{
 const {fixturesRoot}=await import('../../scripts/fixture-inputs.ts'),{join}=await import('node:path'),{selectSharedScenarios,executeAcceptance}=await import('../../scripts/gherkin.ts'),{bindings}=await import('../acceptance/package-comparison.ts');
 const path='workflows/package/semantic-diff.feature',f=selectSharedScenarios(path,await Bun.file(join(fixturesRoot(),path)).text(),['@id-package-diff-equivalent-xml-and-binary-changes']),count=(n:number)=>({implemented:n,planned:0,total:n});
 const result=await executeAcceptance({root:'.',features:[f],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(7)}},bindings,'package-comparison-unit');expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(1);expect(result.counts.steps.passed).toBe(7);
});
test('duplicate member names cannot collapse into a misleading set report',()=>{
 const good=archive([['a.xml','<a/>']]),bad=archive([['a.xml','<a/>'],['a.xml','<b/>']]);
 expect(()=>comparePackageArchives(bad,good)).toThrow(expect.objectContaining({code:'zip-duplicate-entry'}));expect(()=>comparePackageArchives(good,bad)).toThrow(expect.objectContaining({code:'zip-duplicate-entry'}));
});

test('prefix-only XML goes into the new equivalence list but stays changed in OPC byte diff',async()=>{
 const d=Document.create();d.addParagraph('Keep');const original=d.package.toBytes(),parts=readZip(original),source=new TextDecoder().decode(parts.get('word/document.xml')!);
 parts.set('word/document.xml',utf8(source.replaceAll('w:', 'q:').replace('xmlns:w=', 'xmlns:q=')));const modified=writeZip(parts);
 expect(comparePackageArchives(original,modified).equivalent_xml).toEqual(['word/document.xml']);
 expect(diffPackages(await OpcPackage.open(original),await OpcPackage.open(modified)).changed).toEqual(['word/document.xml']);
});
