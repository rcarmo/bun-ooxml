import {describe,expect,test} from 'bun:test';
import { fixturePath, F } from '../../scripts/fixture-inputs.ts';
import {OpcPackage} from '../../src/opc/package.ts';
import {addPart,removePart,addRelationship,removeRelationship,nextPartName,walkParts} from '../../src/opc/graph.ts';
import {getContentType} from '../../src/opc/content-types.ts';
import {parseXml,applyEdits} from '../../src/xml/index.ts';
const fixture=fixturePath(F.shared.presentPlaceholder);
const open=()=>OpcPackage.open(fixture);
describe('guarded OPC graph edits',()=>{
 test('opaque parts and content types survive related save/reopen',async()=>{
  const p=await open(),original=p.get('word/document.xml')!;addPart(p,'custom/data.bin',new Uint8Array([1,2,3]),'application/octet-stream');
  const rel=addRelationship(p,'','urn:test/data','custom/data.bin');const reopened=await OpcPackage.open(p.toBytes());
  expect(reopened.relationships().find(r=>r.id===rel.id)?.resolved).toBe('custom/data.bin');expect(getContentType(reopened,'custom/data.bin')).toBe('application/octet-stream');expect(reopened.get('word/document.xml')).toEqual(original);
 });
 test('collision and missing relationship targets leave package byte-identical',async()=>{
  const p=await open(),before=p.toBytes();expect(()=>addPart(p,'WORD/document.xml','x','application/xml')).toThrow();expect(p.toBytes()).toEqual(before);
  expect(()=>addRelationship(p,'','urn:test/data','absent.xml')).toThrow();expect(p.toBytes()).toEqual(before);
 });
 test('relationship deduplication and smallest available IDs avoid unnecessary edits',async()=>{
  const p=await open();addPart(p,'custom/a.xml','<a/>','application/xml');const first=addRelationship(p,'','urn:test/a','custom/a.xml');const before=p.toBytes();
  expect(addRelationship(p,'','urn:test/a','/custom/a.xml').id).toBe(first.id);expect(p.toBytes()).toEqual(before);
  removeRelationship(p,'',first.id);expect(addRelationship(p,'','urn:test/a','custom/a.xml').id).toBe(first.id);
 });
 test('equivalent relationship leaves existing overrides and XML encoding byte-identical',async()=>{
  const p=await open();addPart(p,'custom/a.xml','<a/>','application/xml');const first=addRelationship(p,'custom/a.xml','urn:test/url','https://example.invalid/',{external:true});
  // Force an explicit rels override, independent of the existing extension default.
  const ct=p.text('[Content_Types].xml'),root=parseXml(ct).root;const q=root.name.includes(':')?root.name.split(':')[0]+':':'';
  p.set('[Content_Types].xml',applyEdits(ct,[{start:root.closeStart,end:root.closeStart,value:`<${q}Override PartName="/custom/_rels/a.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`}]));
  expect(p.text('[Content_Types].xml')).toContain('/custom/_rels/a.xml.rels');
  const before=p.toBytes();expect(addRelationship(p,'custom/a.xml','urn:test/url','https://example.invalid/',{external:true}).id).toBe(first.id);expect(p.toBytes()).toEqual(before);
 });
 test('dollar replacement tokens in external targets are literal bytes, not JS replacement directives',async()=>{
  const p=await open();addPart(p,'custom/a.xml','<a/>','application/xml');
  const rel=addRelationship(p,'custom/a.xml','urn:test/url','https://example.invalid/$&',{external:true});expect(rel.target).toBe('https://example.invalid/$&');
 });
 test('referenced parts refuse until explicitly detached, then only owned metadata is removed',async()=>{
  const p=await open();addPart(p,'custom/a.bin',new Uint8Array([7]),'application/octet-stream');const r=addRelationship(p,'','urn:test/a','custom/a.bin'),before=p.toBytes();
  expect(()=>removePart(p,'custom/a.bin')).toThrow('referenced');expect(p.toBytes()).toEqual(before);
  removeRelationship(p,'',r.id);removePart(p,'custom/a.bin');const reopened=await OpcPackage.open(p.toBytes());expect(reopened.get('custom/a.bin')).toBeUndefined();expect(reopened.text('[Content_Types].xml')).not.toContain('/custom/a.bin');
 });
 test('relationship referenced from owner XML cannot be removed; unknown binary owners refuse',async()=>{
  const p=await open();addPart(p,'custom/b.xml','<b/>','application/xml');const r=addRelationship(p,'word/document.xml','urn:test/b','../custom/b.xml');
  p.set('word/document.xml',p.text('word/document.xml').replace('<w:body>',`<w:body><w:p xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/relationships" q:embed="${r.id}"/>`));
  const before=p.toBytes();expect(()=>removeRelationship(p,'word/document.xml',r.id)).toThrow('referenced');expect(p.toBytes()).toEqual(before);
  addPart(p,'custom/owner.bin',new Uint8Array([1]),'application/octet-stream');const b=addRelationship(p,'custom/owner.bin','urn:test/b','b.xml');const next=p.toBytes();
  expect(()=>removeRelationship(p,'custom/owner.bin',b.id)).toThrow('binary');expect(p.toBytes()).toEqual(next);
 });
 test('cycle-safe traversal excludes external targets and preserves orphan parts',async()=>{
  const p=await open();for(const name of ['a','b','orphan'])addPart(p,`custom/${name}.xml`,`<${name}/>`,'application/xml');
  addRelationship(p,'','urn:test/a','custom/a.xml');addRelationship(p,'custom/a.xml','urn:test/b','b.xml');addRelationship(p,'custom/b.xml','urn:test/a','a.xml');addRelationship(p,'custom/a.xml','urn:test/url','https://example.invalid/',{external:true});
  const parts=walkParts(p);expect(parts.filter(n=>n==='custom/a.xml').length).toBe(1);expect(parts).toContain('custom/b.xml');expect(parts).not.toContain('custom/orphan.xml');expect(p.get('custom/orphan.xml')).toBeDefined();
 });
 test('part name allocation considers orphan and case-colliding names',async()=>{
  const p=await open();addPart(p,'word/header1.xml','<h/>','application/xml');addPart(p,'WORD/header2.xml','<h/>','application/xml');expect(nextPartName(p,'word/header%d.xml')).toBe('word/header3.xml');expect(()=>nextPartName(p,'missing-template')).toThrow();
 });
 test('removing a part removes its own rels but never cascades into targets',async()=>{
  const p=await open();addPart(p,'custom/a.xml','<a/>','application/xml');addPart(p,'custom/b.xml','<b/>','application/xml');addRelationship(p,'custom/a.xml','urn:test/b','b.xml');
  removePart(p,'custom/a.xml');expect(p.get('custom/_rels/a.xml.rels')).toBeUndefined();expect(p.get('custom/b.xml')).toBeDefined();await expect(OpcPackage.open(p.toBytes())).resolves.toBeInstanceOf(OpcPackage);
 });
});
