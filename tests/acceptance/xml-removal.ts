import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { OoxmlError, XmlSnapshot } from '../../src/index.ts';

type State = { source: string; snapshot: XmlSnapshot; output?: string; error?: unknown };
const state = (c: Record<string, unknown>) => c.state as State;
export const scenarioIds = ['@id-xml-go-element-removal-custody', '@id-xml-go-element-removal-refusal'];
export const bindings: StepBinding[] = [
  { pattern: /^the XML source is (.+)$/, run: (c, source) => {
    assert(source);
    Object.assign(state(c), { source, snapshot: XmlSnapshot.parse(source) });
  } },
  { pattern: /^the XML editor removes the p:a subtree and the p:c element from one parsed snapshot$/, run: c => {
    const s = state(c), a = s.snapshot.elements.filter(t => t.localName === 'a' && t.namespaceURI === 'u'),
      others = s.snapshot.elements.filter(t => t.localName === 'c' && t.namespaceURI === 'u');
    assert.equal(a.length, 1); assert.equal(others.length, 1);
    s.output = s.snapshot.remove([...a, ...others]);
  } },
  { pattern: /^the complete output bytes equal (.+)$/, run: (c, expected) => {
    const s = state(c);
    assert.equal(s.error, undefined); assert.equal(typeof s.output, 'string');
    assert.deepEqual(new TextEncoder().encode(s.output), new TextEncoder().encode(expected!));
  } },
  { pattern: /^a separate empty removal returns the exact original source bytes$/, run: c => {
    const s = state(c);
    assert.deepEqual(new TextEncoder().encode(s.snapshot.remove([])), new TextEncoder().encode(s.source));
  } },
  { pattern: /^a removal batch selects (root|p:a and its nested p:b)$/, run: (c, selection) => {
    const s = state(c), targets = selection === 'root' ? [s.snapshot.elements[0]!] :
      s.snapshot.elements.filter(t => (t.localName === 'a' || t.localName === 'b') && t.namespaceURI === 'u');
    assert.equal(targets.length, selection === 'root' ? 1 : 2);
    try { s.output = s.snapshot.remove(targets); } catch (error) { s.error = error; }
  } },
  { pattern: /^the removal returns an error$/, run: c => {
    const s = state(c);
    assert(s.error instanceof OoxmlError);
    assert(['XML_REMOVAL_ROOT', 'XML_REMOVAL_OVERLAP'].includes(s.error.code));
    assert.equal(s.output, undefined);
    assert.equal(s.snapshot.remove([]), s.source);
  } },
];
