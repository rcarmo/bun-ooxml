import assert from 'node:assert/strict';
import type { AcceptanceStep, StepBinding } from '../../scripts/gherkin.ts';
import { OoxmlError, XmlSnapshot, type XmlContent } from '../../src/index.ts';
import { parseXml, attribute } from '../../src/xml/index.ts';

type MatrixRow = { source: string; elementNS: string; attributeNS: string; output: string; noop: string };
type State = { source: string; snapshot: XmlSnapshot; output?: string; error?: unknown; roots?: string[]; namespaces?: string[]; matrix?: MatrixRow[]; value?: string; text?: string };
const state = (c: Record<string, unknown>) => c.state as State;
const element = (localName: string, namespaceURI = '', children: readonly XmlContent[] = []): XmlContent => ({ name: { localName, namespaceURI }, children });
export const scenarioIds = ['@id-xml-go-child-insertion-custody', '@id-xml-go-child-insertion-refusal', '@id-xml-go-child-namespace-matrix', '@id-xml-go-element-replacement-custody', '@id-xml-go-element-replacement-refusal'];
function table(c: Record<string, unknown>, header: string): string[] {
  const rows = (c.step as AcceptanceStep).argument?.dataTable; assert(rows && rows.length > 1); assert.deepEqual(rows[0], [header]);
  return rows.slice(1).map(row => { assert.equal(row.length, 1); return row[0]!; });
}
export const bindings: StepBinding[] = [
  { pattern: /^a structured edit inserts a new-namespace x child with an other-namespace a attribute and a plain text child under the existing a element$/, run: c => {
    const s = state(c), target = s.snapshot.elements.find(t => t.localName === 'a' && t.namespaceURI === 'u'); assert(target);
    s.output = s.snapshot.appendChildren([{ target, children: [{ name: { localName: 'x', namespaceURI: 'new' }, attributes: [{ name: { localName: 'a', namespaceURI: 'other' }, value: 'value' }], children: [element('plain', '', ['text'])] }] }]);
  } },
  { pattern: /^reparsing finds expanded element names new\/x and empty-namespace plain$/, run: c => {
    const s = state(c); assert(s.output); const root = parseXml(s.output).root, a = root.children[0]!, x = a.children[0]!;
    assert.deepEqual([x.localName, x.namespaceURI], ['x', 'new']); assert.deepEqual([x.children[0]!.localName, x.children[0]!.namespaceURI], ['plain', '']);
  } },
  { pattern: /^the unedited sibling bytes (.+) remain in the output$/, run: (c, sibling) => { const s = state(c); assert(sibling && s.source.includes(sibling)); assert(s.output?.includes(sibling)); } },
  { pattern: /^one structured insertion batch targets both the root and its nested a element$/, run: c => {
    const s = state(c), a = s.snapshot.elements.find(t => t.localName === 'a' && t.namespaceURI === 'u'); assert(a);
    try { s.output = s.snapshot.appendChildren([s.snapshot.elements[0]!, a].map(target => ({ target, children: [element('x')] }))); } catch (error) { s.error = error; }
  } },
  { pattern: /^the insertion returns an error$/, run: c => { const s = state(c); assert(s.error instanceof OoxmlError); assert.equal(s.error.code, 'XML_STRUCTURE_OVERLAP'); assert.equal(s.output, undefined); } },
  { pattern: /^a separate empty insertion batch returns the exact original source bytes$/, run: c => { const s = state(c); assert.equal(s.snapshot.appendChildren([]), s.source); } },
  { pattern: /^these four XML root sources, each interpreted as a JSON string$/, run: c => { state(c).roots = table(c, 'source_json').map(row => { const value: unknown = JSON.parse(row); assert.equal(typeof value, 'string'); return value as string; }); } },
  { pattern: /^these five namespace URI choices independently for each child element and flag attribute$/, run: c => { state(c).namespaces = table(c, 'namespace_uri'); } },
  { pattern: /^the XML editor inserts child with a flag attribute of JSON value (.+) and a plain grandchild of JSON text (.+) for all 4 by 5 by 5 choices$/, run: (c, value, text) => {
    const s = state(c); assert.equal(s.roots?.length, 4); assert.equal(s.namespaces?.length, 5);
    s.value = JSON.parse(value!); s.text = JSON.parse(text!); assert.equal(typeof s.value, 'string'); assert.equal(typeof s.text, 'string'); s.matrix = [];
    for (const source of s.roots!) for (const elementNS of s.namespaces!) for (const attributeNS of s.namespaces!) {
      const snapshot = XmlSnapshot.parse(source);
      const output = snapshot.appendChildren([{ target: snapshot.elements[0]!, children: [{ name: { localName: 'child', namespaceURI: elementNS }, attributes: [{ name: { localName: 'flag', namespaceURI: attributeNS }, value: s.value! }], children: [element('plain', '', [s.text!])] }] }]);
      s.matrix.push({ source, elementNS, attributeNS, output, noop: snapshot.appendChildren([]) });
    }
  } },
  { pattern: /^reparsing preserves the root child and grandchild expanded names and the child attribute name and value for every choice$/, run: c => {
    const s = state(c); assert.equal(s.matrix?.length, 100);
    for (const row of s.matrix!) {
      const original = parseXml(row.source).root, root = parseXml(row.output).root; assert.equal(root.children.length, 1);
      const child = root.children[0]!; assert.equal(child.children.length, 1); const grandchild = child.children[0]!;
      assert.deepEqual([root.localName, root.namespaceURI], [original.localName, original.namespaceURI]);
      assert.deepEqual([child.localName, child.namespaceURI], ['child', row.elementNS]);
      assert.deepEqual([grandchild.localName, grandchild.namespaceURI], ['plain', '']);
      assert.equal(attribute(child, 'flag', row.attributeNS), s.value);
    }
  } },
  { pattern: /^the grandchild text equals JSON (.+) for every choice$/, run: (c, value) => {
    const s = state(c); assert.equal(s.matrix?.length, 100); const expected: unknown = JSON.parse(value!); assert.equal(typeof expected, 'string');
    for (const row of s.matrix!) assert.equal(parseXml(row.output).root.children[0]!.children[0]!.text, expected);
  } },
  { pattern: /^a separate empty edit returns each exact original root source$/, run: c => { const s = state(c); assert.equal(s.matrix?.length, 100); for (const row of s.matrix!) assert.equal(row.noop, row.source); } },
  { pattern: /^the XML editor replaces p:old with a bound-namespace new element containing value and an empty-namespace plain element$/, run: c => {
    const s = state(c), target = s.snapshot.elements.find(t => t.localName === 'old' && t.namespaceURI === 'inner'); assert(target);
    s.output = s.snapshot.replaceElements([{ target, children: [element('new', 'bound', ['value']), element('plain')] }]);
  } },
  { pattern: /^a replacement batch selects (root|p:old twice|p:old and its nested p:child)$/, run: (c, selection) => {
    const s = state(c), old = s.snapshot.elements.find(t => t.localName === 'old'), child = s.snapshot.elements.find(t => t.localName === 'child'); assert(old && child);
    const targets = selection === 'root' ? [s.snapshot.elements[0]!] : selection === 'p:old twice' ? [old, old] : [old, child];
    try { s.output = s.snapshot.replaceElements(targets.map(target => ({ target, children: [element('new')] }))); } catch (error) { s.error = error; }
  } },
  { pattern: /^replacement returns an error and no edited output$/, run: c => { const s = state(c); assert(s.error instanceof OoxmlError); assert(['XML_STRUCTURE_ROOT', 'XML_STRUCTURE_OVERLAP'].includes(s.error.code)); assert.equal(s.output, undefined); } },
  { pattern: /^a separate empty replacement returns the exact original source bytes$/, run: c => { const s = state(c); assert.equal(s.snapshot.replaceElements([]), s.source); } },
];
