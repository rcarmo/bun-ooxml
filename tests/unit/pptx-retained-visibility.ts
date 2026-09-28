import assert from 'node:assert/strict';
import { OpcPackage } from '../../src/opc/package.ts';
import { Presentation } from '../../src/pptx/index.ts';
import { attribute, elements, parseXml } from '../../src/xml/index.ts';

const presentationNs = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const officeRelNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const slideRel = officeRelNs + '/slide';
const expectedParts = [1, 2, 3, 4].map(n => `ppt/slides/slide${n}.xml`);
const expectedIds = ['256', '257', '258', '259'];

/** Package-only observation; the originals have not been verified in PowerPoint. */
export async function inspectRetainedVisibility(bytes: Uint8Array, role: 'source' | 'control') {
  const pkg = await OpcPackage.open(bytes);
  const presentation = await Presentation.open(bytes);
  const mainPart = pkg.mainPart();
  assert.equal(mainPart, 'ppt/presentation.xml');
  const root = parseXml(pkg.text(mainPart)).root;
  assert.equal(root.localName, 'presentation'); assert.equal(root.namespaceURI, presentationNs);
  const lists = elements(root, 'sldIdLst', presentationNs);
  assert.equal(lists.length, 1);
  const slideIds = elements(lists[0]!, 'sldId', presentationNs);
  assert.equal(slideIds.length, 4);
  assert.equal(lists[0]!.children.length, 4);
  const relationships = pkg.relationships(mainPart);
  const byId = new Map(relationships.map(r => [r.id, r]));
  assert.equal(byId.size, relationships.length);
  const ids: string[] = [], relationIds: string[] = [], parts: string[] = [];
  for (const [index, entry] of slideIds.entries()) {
    const id = attribute(entry, 'id'), rid = attribute(entry, 'id', officeRelNs);
    ids.push(id ?? ''); relationIds.push(rid ?? '');
    assert.equal(id, expectedIds[index]);
    assert.equal(rid, role === 'source' ? `rId${7 + index}` : `rId${2 + index}`);
    const relation = byId.get(rid ?? '');
    assert(relation && !relation.external && relation.type === slideRel);
    const part = relation.resolved;
    assert(part); assert.equal(part, expectedParts[index]);
    assert(pkg.names().includes(part)); parts.push(part);
    const slide = parseXml(pkg.text(part)).root;
    assert.equal(slide.localName, 'sld'); assert.equal(slide.namespaceURI, presentationNs);
    assert.equal(attribute(slide, 'show'), undefined);
    assert.equal(attribute(slide, 'show', presentationNs), role === 'source' && index === 2 ? '0' : undefined);
    assert.equal(Object.keys(slide.attributes).filter(name => name === 'show' || name.endsWith(':show')).length, role === 'source' && index === 2 ? 1 : 0);
  }
  assert.equal(new Set(ids).size, 4); assert.equal(new Set(relationIds).size, 4); assert.equal(new Set(parts).size, 4);
  assert.deepEqual(presentation.slides.map(slide => slide.partName), parts);
  assert.deepEqual(pkg.toBytes(), bytes);
  return { slideIds: ids, relationships: relationIds, parts };
}
