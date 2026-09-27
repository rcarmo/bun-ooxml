import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { OoxmlError, type XmlSnapshot } from '../../src/index.ts';

type State = { source: string; snapshot: XmlSnapshot; output?: string; error?: unknown };
const state = (c: Record<string, unknown>) => c.state as State;
function target(s: State) { const selected = s.snapshot.elements.find(e => e.localName === 't' && e.namespaceURI === ''); assert(selected); return selected; }
export const scenarioIds = ['@id-xml-go-attribute-splice-custody', '@id-xml-go-attribute-batch-refusal'];
export const bindings: StepBinding[] = [
  { pattern: /^a lexical edit sets the attribute (\S+) of the first t element to (.+)$/, run: (c, name, value) => {
    const s = state(c); assert(name && value);
    s.output = s.snapshot.setAttributes([{ target: target(s), name, value }]);
    assert.equal(s.snapshot.setAttributes([]), s.source);
  } },
  { pattern: /^one lexical edit batch sets a of the first t element to x and to y$/, run: c => {
    const s = state(c), selected = target(s);
    try { s.output = s.snapshot.setAttributes([{ target: selected, name: 'a', value: 'x' }, { target: selected, name: 'a', value: 'y' }]); }
    catch (error) { s.error = error; }
  } },
  { pattern: /^the edit returns an error instead of accepting that batch$/, run: c => {
    const s = state(c); assert(s.error instanceof OoxmlError); assert.equal(s.error.code, 'XML_ATTRIBUTE_DUPLICATE');
    assert.equal(s.output, undefined); assert.equal(s.snapshot.setAttributes([]), s.source);
  } },
];
