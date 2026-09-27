import { expect, test } from 'bun:test';
import { XmlSnapshot, type XmlContent } from '../../src/index.ts';
import { parseXml, attribute } from '../../src/xml/index.ts';
const element = (localName: string, namespaceURI = '', children: readonly XmlContent[] = []): XmlContent => ({ name: { localName, namespaceURI }, children });
const code = (code: string) => expect.objectContaining({ code });

test('structured insertion preserves parent tag spelling, tail, comments and siblings', () => {
  const source = '<root xmlns="u" xmlns:n1="occupied"><a x = \'keep\' /><b>keep</b></root>';
  const s = XmlSnapshot.parse(source);
  const output = s.appendChildren([{ target: s.elements[1]!, children: [{ name: { localName: 'x', namespaceURI: 'new' }, attributes: [{ name: { localName: 'a', namespaceURI: 'other' }, value: 'value' }], children: [element('plain', '', ['text'])] }] }]);
  const d = parseXml(output), a = d.root.children[0]!, x = a.children[0]!;
  expect(output).toStartWith('<root xmlns="u" xmlns:n1="occupied"><a x = \'keep\' >');
  expect(output).toEndWith('</a><b>keep</b></root>');
  expect([x.localName, x.namespaceURI]).toEqual(['x', 'new']);
  expect(attribute(x, 'a', 'other')).toBe('value');
  expect([x.children[0]!.localName, x.children[0]!.namespaceURI, x.children[0]!.text]).toEqual(['plain', '', 'text']);
  expect(s.appendChildren([])).toBe(source);
});

test('replacement renders only in the surviving parent scope with exact canonical custody', () => {
  const source = '<root xmlns="outer" xmlns:p="bound"><!--a--><p:old xmlns:p="inner" x=\'1\'><p:child/></p:old> tail <last/></root>';
  const s = XmlSnapshot.parse(source);
  const output = s.replaceElements([{ target: s.elements[1]!, children: [element('new', 'bound', ['value']), element('plain')] }]);
  expect(output).toBe('<root xmlns="outer" xmlns:p="bound"><!--a--><p:new>value</p:new><plain xmlns=""/> tail <last/></root>');
  expect(s.replaceElements([])).toBe(source);
  expect(s.remove([])).toBe(source);
});

test('4 by 5 by 5 namespace matrix preserves element, attribute and text meanings', () => {
  const roots = ['<r/>', '<r xmlns="u"/>', '<p:r xmlns:p="u"/>', '<r xmlns="u" xmlns:n1="v" xmlns:n2="occupied"/>'];
  const namespaces = ['', 'u', 'v', 'fresh', 'http://www.w3.org/XML/1998/namespace'];
  for (const source of roots) for (const ens of namespaces) for (const ans of namespaces) {
    const s = XmlSnapshot.parse(source), rootBefore = parseXml(source).root;
    const value = '\t\r\n & 😀', text = 'x\ry\nz';
    const output = s.appendChildren([{ target: s.elements[0]!, children: [{ name: { localName: 'child', namespaceURI: ens }, attributes: [{ name: { localName: 'flag', namespaceURI: ans }, value }], children: [element('plain', '', [text])] }] }]);
    const root = parseXml(output).root, child = root.children[0]!, grandchild = child.children[0]!;
    expect([root.localName, root.namespaceURI]).toEqual([rootBefore.localName, rootBefore.namespaceURI]);
    expect([child.localName, child.namespaceURI]).toEqual(['child', ens]);
    expect(attribute(child, 'flag', ans)).toBe(value);
    expect([grandchild.localName, grandchild.namespaceURI, grandchild.text]).toEqual(['plain', '', text]);
    expect(s.appendChildren([])).toBe(source);
  }
});

test('structured batches reject root replacement, overlapping and duplicate targets including empty payloads', () => {
  const source = '<r><a><b/></a><c/></r>', s = XmlSnapshot.parse(source);
  expect(() => s.replaceElements([{ target: s.elements[0]!, children: [] }])).toThrow(code('XML_STRUCTURE_ROOT'));
  for (const method of ['appendChildren', 'replaceElements'] as const) {
    for (const indexes of [[1, 1], [1, 2], [2, 1]]) {
      expect(() => s[method](indexes.map(i => ({ target: s.elements[i]!, children: [] })))).toThrow(code('XML_STRUCTURE_OVERLAP'));
      expect(s[method]([])).toBe(source);
    }
  }
});

test('disjoint replacements and insertions preserve original order independent of target order', () => {
  const source = '<r>😀<a/>\r\n<!--gap--><b>old</b>tail</r>', s = XmlSnapshot.parse(source);
  expect(s.appendChildren([{ target: s.elements[2]!, children: [' & more'] }, { target: s.elements[1]!, children: [element('c')] }])).toBe('<r>😀<a><c/></a>\r\n<!--gap--><b>old &amp; more</b>tail</r>');
  expect(s.replaceElements([{ target: s.elements[2]!, children: [] }, { target: s.elements[1]!, children: [element('c')] }])).toBe('<r>😀<c/>\r\n<!--gap-->tail</r>');
  expect(s.appendChildren([{ target: s.elements[1]!, children: [] }])).toBe(source);
});

test('new child scopes do not leak into siblings or reuse shadowed parent prefixes', () => {
  const s = XmlSnapshot.parse('<r xmlns:p="u"><a xmlns:p="v"/><b/></r>');
  const out = s.appendChildren([{ target: s.elements[1]!, children: [element('x', 'u'), element('x', 'v'), element('x', '')] }]);
  const nodes = parseXml(out).root.children[0]!.children;
  expect(nodes.map(n => n.namespaceURI)).toEqual(['u', 'v', '']);
  expect(out).toContain('<p:x/>'); expect(out).toEndWith('</a><b/></r>');
});

test('structured edits require issued handles, frozen inputs are accepted and caller iterators ignored', () => {
  const s = XmlSnapshot.parse('<r><a/></r>'), foreign = XmlSnapshot.parse('<r><a/></r>');
  for (const target of [foreign.elements[1], { ...s.elements[1]! }, null]) expect(() => s.appendChildren([{ target: target as never, children: [] }])).toThrow(code('XML_STRUCTURE_TARGET'));
  const children = [element('x')]; children[Symbol.iterator] = function* () { throw Error('do not iterate'); };
  const patches = [{ target: s.elements[1]!, children }]; patches[Symbol.iterator] = function* () { throw Error('do not iterate'); };
  expect(s.appendChildren(Object.freeze(patches))).toBe('<r><a><x/></a></r>');
  expect(() => XmlSnapshot.parse('<r><a/></r>').replaceElements(patches)).toThrow(code('XML_STRUCTURE_TARGET'));
});

test('invalid structured names, namespace declarations, duplicate attributes and data refuse atomically', () => {
  const s = XmlSnapshot.parse('<r><a/></r>'), target = s.elements[1]!;
  const invalid: unknown[] = [element('p:x'), element('1bad'), element('x', 'http://www.w3.org/2000/xmlns/'),
    { name: { localName: 'x', namespaceURI: '' }, attributes: [{ name: { localName: 'xmlns', namespaceURI: '' }, value: 'u' }] },
    { name: { localName: 'x', namespaceURI: '' }, attributes: ['a', 'a'].map(localName => ({ name: { localName, namespaceURI: 'u' }, value: 'v' })) },
    element('x', '', ['\ud800']), element('x', '', ['\u0000']), null, 12, { name: { localName: 'x', namespaceURI: '' }, children: [undefined] }];
  for (const child of invalid) {
    let output: string | undefined;
    expect(() => { output = s.appendChildren([{ target, children: [element('good'), child as XmlContent] }]); }).toThrow();
    expect(output).toBeUndefined(); expect(s.appendChildren([])).toBe('<r><a/></r>');
  }
});

test('bounded renderer refuses cycles, excessive depth, node counts and expanded output', () => {
  const s = XmlSnapshot.parse('<r/>'), target = s.elements[0]!;
  const cycle: any = { name: { localName: 'x', namespaceURI: '' }, children: [] }; cycle.children.push(cycle);
  expect(() => s.appendChildren([{ target, children: [cycle] }])).toThrow(code('XML_STRUCTURE_CYCLE'));
  let nested = element('x'); for (let i = 0; i < 256; i++) nested = element('x', '', [nested]);
  for (const children of [[nested], new Array(100001).fill(element('x')), ['&'.repeat(2 * 1024 * 1024)]]) {
    expect(() => s.appendChildren([{ target, children }])).toThrow(code('XML_STRUCTURE_LIMIT'));
  }
  expect(() => s.appendChildren(new Array(100001).fill({ target, children: [] }))).toThrow(code('XML_STRUCTURE_LIMIT'));
});

test('empty replacements validate newly adjacent text before returning output', () => {
  const xml = '<r>]]<a/>>tail</r>', s = XmlSnapshot.parse(xml);
  expect(() => s.replaceElements([{ target: s.elements[1]!, children: [] }])).toThrow(code('XML_STRUCTURE_UNSAFE'));
  expect(s.replaceElements([])).toBe(xml);
  expect(s.replaceElements([{ target: s.elements[1]!, children: ['>'] }])).toBe('<r>]]&gt;>tail</r>');
});

test('generated prefixes avoid occupied bindings and scale across many distinct attribute namespaces', () => {
  const s = XmlSnapshot.parse('<r xmlns:n1="occupied" xmlns:n2="also"/>');
  const attributes = Array.from({ length: 1500 }, (_, i) => ({ name: { localName: 'flag', namespaceURI: `urn:${i}` }, value: String(i) }));
  const output = s.appendChildren([{ target: s.elements[0]!, children: [{ name: { localName: 'child', namespaceURI: 'new' }, attributes }] }]);
  expect(output).toContain('<n3:child');
  const child = parseXml(output).root.children[0]!;
  expect(attribute(child, 'flag', 'urn:0')).toBe('0'); expect(attribute(child, 'flag', 'urn:1499')).toBe('1499');
  expect(child.namespaceURI).toBe('new');
});

test('rendered size budget includes inherited-prefix names, declarations and self-closing expansion', () => {
  const limit = 8 * 1024 * 1024, source = '<r/>', s = XmlSnapshot.parse(source);
  // <r> + text + </r> has seven code units of markup.
  const text = 'a'.repeat(limit - 7);
  expect(s.appendChildren([{ target: s.elements[0]!, children: [text] }]).length).toBe(limit);
  expect(() => s.appendChildren([{ target: s.elements[0]!, children: [text + 'a'] }])).toThrow(code('XML_STRUCTURE_LIMIT'));
  expect(() => s.appendChildren([{ target: s.elements[0]!, children: [element('x', '"'.repeat(limit / 2))] }])).toThrow(code('XML_STRUCTURE_LIMIT'));
});

test('depth limits include existing ancestors and shared authored objects are independent siblings', () => {
  const xml = '<r>'.repeat(255) + '<a/>' + '</r>'.repeat(255), s = XmlSnapshot.parse(xml);
  expect(() => s.appendChildren([{ target: s.elements[255]!, children: [element('x')] }])).toThrow(code('XML_STRUCTURE_LIMIT'));
  const shared = element('x', 'u', ['<&\r']); const root = XmlSnapshot.parse('<r/>');
  const output = root.appendChildren([{ target: root.elements[0]!, children: [shared, shared] }]);
  expect(parseXml(output).root.children.map(n => [n.localName, n.namespaceURI, n.text])).toEqual([['x', 'u', '<&\r'], ['x', 'u', '<&\r']]);
});

test('seven canonical structured-edit cases execute without activating the byte-input seed case', async () => {
  const { fixturesRoot } = await import('../../scripts/fixture-inputs.ts'), { join } = await import('node:path');
  const { executeAcceptance, selectSharedScenarios } = await import('../../scripts/gherkin.ts');
  const { bindings } = await import('../acceptance/steps.ts'), { scenarioIds } = await import('../acceptance/xml-structure.ts');
  const path = 'workflows/xml/editing.feature', text = await Bun.file(join(fixturesRoot(), path)).text();
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  const inventory = (source: string) => ({ root: '.', features: [selectSharedScenarios(path, source, scenarioIds)], counts: { features: count(1), scenarios: count(5), cases: count(7), steps: count(30) } });
  const good = await executeAcceptance(inventory(text), bindings, 'xml-structure-unit');
  expect(good.failures).toEqual([]); expect(good.counts.cases.passed).toBe(7); expect(good.counts.cases.planned).toBe(9);
  const bad = await executeAcceptance(inventory(text.replace('the grandchild text equals JSON "x\\ry\\nz"', 'the grandchild text equals JSON "wrong"')), bindings, 'xml-structure-value-control');
  expect(bad.counts.cases.failed).toBe(1); expect(bad.counts.steps.failed).toBe(1); expect(bad.counts.steps.undefined).toBe(0); expect(bad.counts.steps.ambiguous).toBe(0);
});

test('structured acceptance predicates reject corrupted namespaces, text, no-ops and fabricated refusals', async () => {
  const { fixturesRoot } = await import('../../scripts/fixture-inputs.ts'), { join } = await import('node:path');
  const { executeAcceptance, selectSharedScenarios } = await import('../../scripts/gherkin.ts');
  const { bindings } = await import('../acceptance/steps.ts'), { scenarioIds } = await import('../acceptance/xml-structure.ts');
  const path = 'workflows/xml/editing.feature', text = await Bun.file(join(fixturesRoot(), path)).text();
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  type State = { output?: string; error?: unknown; matrix?: { output: string; noop: string }[] };
  const matrixStep = 'the XML editor inserts child with a flag attribute of JSON value "\\t\\r\\n & 😀" and a plain grandchild of JSON text "x\\ry\\nz" for all 4 by 5 by 5 choices';
  const mutations: [string, (s: State) => void][] = [
    [matrixStep, s => { s.matrix![0]!.output = '<r><wrong flag="v"><plain/></wrong></r>'; }],
    [matrixStep, s => { s.matrix![0]!.noop = 'wrong'; }],
    ['a replacement batch selects root', s => { s.output = '<r/>'; s.error = undefined; }],
    ['one structured insertion batch targets both the root and its nested a element', s => { s.error = new Error('untyped'); }],
    ['the XML editor replaces p:old with a bound-namespace new element containing value and an empty-namespace plain element', s => { s.output = '<root/>'; }],
  ];
  for (const [step, mutate] of mutations) {
    const corrupted = bindings.map(b => b.pattern.test(step) ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => { await b.run(c, ...captures); mutate(c.state as State); } } : b);
    const result = await executeAcceptance({ root: '.', features: [selectSharedScenarios(path, text, scenarioIds)], counts: { features: count(1), scenarios: count(5), cases: count(7), steps: count(30) } }, corrupted, 'structure-controls');
    expect(result.counts.cases.failed).toBeGreaterThan(0); expect(result.counts.steps.failed).toBeGreaterThan(0);
    expect(result.counts.steps.undefined).toBe(0); expect(result.counts.steps.ambiguous).toBe(0);
  }
});

test('namespace scope work has a separate bounded refusal for repeated broad scopes', () => {
  const source = '<r ' + Array.from({ length: 1500 }, (_, i) => `xmlns:p${i}="urn:${i}"`).join(' ') + '/>';
  const s = XmlSnapshot.parse(source);
  expect(() => s.appendChildren([{ target: s.elements[0]!, children: new Array(700).fill(element('child')) }])).toThrow(code('XML_STRUCTURE_LIMIT'));
  expect(s.appendChildren([])).toBe(source);
  expect(parseXml(s.appendChildren([{ target: s.elements[0]!, children: [element('child')] }])).root.children[0]!.localName).toBe('child');
});
