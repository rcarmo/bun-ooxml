import { expect, test } from 'bun:test';
import { OoxmlError, XmlSnapshot } from '../../src/index.ts';
import { parseXml } from '../../src/xml/index.ts';

const source = "<r xmlns:p=\"u\"><!--keep--><p:a x = '1'><p:b>text</p:b></p:a> gap <p:c /></r>";
const code = (value: string) => expect.objectContaining({ code: value });

test('removing disjoint XML elements preserves all surrounding source characters', () => {
  const snapshot = XmlSnapshot.parse(source);
  const [, a, , c] = snapshot.elements;
  expect(snapshot.remove([c!, a!])).toBe('<r xmlns:p="u"><!--keep--> gap </r>');
  expect(snapshot.remove([])).toBe(source);
  expect(snapshot.remove([a!])).toBe('<r xmlns:p="u"><!--keep--> gap <p:c /></r>');
  expect(snapshot.remove([c!, a!])).toBe('<r xmlns:p="u"><!--keep--> gap </r>');
});

test('root, duplicate and ancestor-descendant removal refuse atomically', () => {
  const snapshot = XmlSnapshot.parse(source);
  const [root, a, b] = snapshot.elements;
  for (const [targets, error] of [
    [[root!], 'XML_REMOVAL_ROOT'], [[a!, a!], 'XML_REMOVAL_OVERLAP'],
    [[a!, b!], 'XML_REMOVAL_OVERLAP'], [[b!, a!], 'XML_REMOVAL_OVERLAP'],
  ] as const) {
    expect(() => snapshot.remove(targets)).toThrow(code(error));
    expect(snapshot.remove([])).toBe(source);
  }
  expect(snapshot.remove([b!])).toBe(source.replace('<p:b>text</p:b>', ''));
});

test('handles belong only to their issuing snapshot, even for identical source', () => {
  const snapshot = XmlSnapshot.parse(source), other = XmlSnapshot.parse(source);
  const a = snapshot.elements[1]!;
  for (const target of [other.elements[1], { ...a }, { ...a, start: 0 }, null, undefined, 1]) {
    expect(() => snapshot.remove([target as typeof a])).toThrow(code('XML_REMOVAL_TARGET'));
  }
  expect(() => snapshot.remove(null as never)).toThrow(code('XML_REMOVAL_TARGET'));
  expect(snapshot.remove([])).toBe(source);
  const result = snapshot.remove([a]);
  expect(() => XmlSnapshot.parse(result).remove([a])).toThrow(code('XML_REMOVAL_TARGET'));
});

test('snapshots and target offsets are immutable and names use expanded namespaces', () => {
  const snapshot = XmlSnapshot.parse('<r xmlns="u"><p:a xmlns:p="v"/><a xmlns=""/></r>');
  expect(Object.isFrozen(snapshot)).toBe(true);
  expect(Object.isFrozen(snapshot.elements)).toBe(true);
  expect(snapshot.elements.every(Object.isFrozen)).toBe(true);
  expect(snapshot.elements.map(({ name, localName, namespaceURI }) => ({ name, localName, namespaceURI }))).toEqual([
    { name: 'r', localName: 'r', namespaceURI: 'u' },
    { name: 'p:a', localName: 'a', namespaceURI: 'v' },
    { name: 'a', localName: 'a', namespaceURI: '' },
  ]);
  expect(() => { (snapshot.elements[1] as { start: number }).start = 0; }).toThrow();
  expect(snapshot.remove([snapshot.elements[1]!])).toBe('<r xmlns="u"><a xmlns=""/></r>');
});

test('UTF-16 offsets preserve astral characters, raw line endings, declarations and PIs', () => {
  const xml = '<?xml version="1.0"?>\r\n<?keep v?>\r\n<r>😀\r\n<x a="&amp;">雪</x><!--tail-->\r\n</r>\r\n';
  const snapshot = XmlSnapshot.parse(xml), x = snapshot.elements[1]!;
  expect(xml.slice(x.start, x.end)).toBe('<x a="&amp;">雪</x>');
  expect(snapshot.remove([x])).toBe(xml.replace('<x a="&amp;">雪</x>', ''));
  expect(parseXml(snapshot.remove([x])).root.text).toBe('😀\n\n');
});

test('removal validates newly adjacent text before returning output', () => {
  const xml = '<r>]]<a/>&gt;<b/>>tail</r>', snapshot = XmlSnapshot.parse(xml);
  expect(snapshot.remove([snapshot.elements[1]!])).toBe('<r>]]&gt;<b/>>tail</r>');
  const dangerous = XmlSnapshot.parse('<r>]]<a/>>tail</r>');
  expect(() => dangerous.remove([dangerous.elements[1]!])).toThrow(code('XML_REMOVAL_UNSAFE'));
  expect(dangerous.remove([])).toBe('<r>]]<a/>>tail</r>');
});

test('snapshot construction refuses malformed, DTD-bearing and oversized input', () => {
  for (const xml of ['<r>', '<!DOCTYPE r><r/>', '<r>\u0000</r>', '<r>' + 'x'.repeat(8 * 1024 * 1024) + '</r>', '<r>'.repeat(257) + '</r>'.repeat(257)]) {
    expect(() => XmlSnapshot.parse(xml)).toThrow(OoxmlError);
  }
  for (const input of [null, undefined, 1, {}]) expect(() => XmlSnapshot.parse(input as string)).toThrow(code('XML_REMOVAL_SOURCE'));
});

test('removal preserves unselected descendants and never mutates selection arrays', () => {
  const snapshot = XmlSnapshot.parse('<r><a/><b><c/></b><d/></r>');
  const targets = Object.freeze([snapshot.elements[4]!, snapshot.elements[1]!]);
  expect(snapshot.remove(targets)).toBe('<r><b><c/></b></r>');
  expect(targets.map(t => t.name)).toEqual(['d', 'a']);
  expect(snapshot.elements.map(t => t.name)).toEqual(['r', 'a', 'b', 'c', 'd']);
  expect(XmlSnapshot.parse('<r/>').remove([])).toBe('<r/>');
});

test('bounded snapshots do not collect descendant text at every ancestor', () => {
  const prefix = '<r>'.repeat(200), suffix = '</r>'.repeat(200);
  const xml = prefix + 'x'.repeat(256 * 1024) + '<drop/>' + suffix;
  const snapshot = XmlSnapshot.parse(xml);
  expect(snapshot.elements).toHaveLength(201);
  expect(snapshot.elements.every(t => !('text' in t) && !('children' in t))).toBe(true);
  expect(snapshot.remove([snapshot.elements[200]!])).toBe(prefix + 'x'.repeat(256 * 1024) + suffix);
});

test('target selection uses bounded array entries rather than an overridable iterator', () => {
  const snapshot = XmlSnapshot.parse('<r><a/><b/></r>');
  const targets = [snapshot.elements[1]!];
  targets[Symbol.iterator] = function* () { yield snapshot.elements[2]!; return undefined; };
  expect(snapshot.remove(targets)).toBe('<r><b/></r>');
});

test('three canonical XML removal cases execute while unrelated shared operations stay planned', async () => {
  const { fixturesRoot } = await import('../../scripts/fixture-inputs.ts');
  const { join } = await import('node:path');
  const { selectSharedScenarios, executeAcceptance } = await import('../../scripts/gherkin.ts');
  const { bindings, scenarioIds } = await import('../acceptance/xml-removal.ts');
  const path = 'workflows/xml/parsing.feature';
  const feature = selectSharedScenarios(path, await Bun.file(join(fixturesRoot(), path)).text(), scenarioIds);
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  const result = await executeAcceptance({ root: '.', features: [feature], counts: {
    features: count(1), scenarios: count(2), cases: count(3), steps: count(10),
  } }, bindings, 'xml-removal-unit');
  expect(result.failures).toEqual([]);
  expect(result.counts.cases.passed).toBe(3);
  expect(result.counts.cases.planned).toBeGreaterThan(0);
});
