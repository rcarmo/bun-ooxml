import { expect, test } from 'bun:test';
import { XmlSnapshot, OoxmlError } from '../../src/index.ts';
import { attribute, parseXml } from '../../src/xml/index.ts';
const source = '<r xmlns:p="urn:p"><t a = \'a&amp;b\' p:n="old"/><t>text</t></r>';
const code = (code: string) => expect.objectContaining({ code });

test('attribute replacement preserves original quote, whitespace and unrelated source', () => {
  const s = XmlSnapshot.parse(source), target = s.elements[1]!;
  expect(s.setAttributes([{ target, name: 'a', value: "x'y&z" }])).toBe('<r xmlns:p="urn:p"><t a = \'x&#39;y&amp;z\' p:n="old"/><t>text</t></r>');
  expect(s.setAttributes([{ target, name: 'p:n', value: 'new' }])).toBe('<r xmlns:p="urn:p"><t a = \'a&amp;b\' p:n="new"/><t>text</t></r>');
  expect(s.setAttributes([{ target, name: 'fresh', value: 'value' }])).toBe('<r xmlns:p="urn:p"><t a = \'a&amp;b\' p:n="old" fresh="value"/><t>text</t></r>');
  expect(s.setAttributes([])).toBe(source);
});

test('same decoded attribute value is a lexical no-op and edits never consume the snapshot', () => {
  const s = XmlSnapshot.parse('<r a="&#65;&apos;"/>'), target = s.elements[0]!;
  expect(s.setAttributes([{ target, name: 'a', value: "A'" }])).toBe('<r a="&#65;&apos;"/>');
  expect(s.setAttributes([{ target, name: 'a', value: 'next' }])).toBe('<r a="next"/>');
  expect(s.setAttributes([])).toBe('<r a="&#65;&apos;"/>');
  expect(s.setAttributes([{ target, name: 'a', value: 'next' }])).toBe('<r a="next"/>');
});

test('attribute names resolve in target scope, aliases edit existing lexical name and default namespaces do not apply', () => {
  const xml = '<r xmlns="u" xmlns:p="v" xmlns:q="v"><c q:a="old" a="plain"/><d xmlns:p="w" p:a="inner"/></r>';
  const s = XmlSnapshot.parse(xml), c = s.elements[1]!, d = s.elements[2]!;
  const output = s.setAttributes([{ target: c, name: 'p:a', value: 'changed' }, { target: c, name: 'a', value: 'new' }, { target: d, name: 'p:a', value: 'local' }, { target: c, name: 'xml:lang', value: 'en' }]);
  expect(output).toBe('<r xmlns="u" xmlns:p="v" xmlns:q="v"><c q:a="changed" a="new" xml:lang="en"/><d xmlns:p="w" p:a="local"/></r>');
  const doc = parseXml(output); expect(attribute(doc.root.children[0]!, 'a', 'v')).toBe('changed');
  expect(attribute(doc.root.children[0]!, 'a')).toBe('new'); expect(attribute(doc.root.children[1]!, 'a', 'w')).toBe('local');
  expect(attribute(doc.root.children[0]!, 'lang', 'http://www.w3.org/XML/1998/namespace')).toBe('en');
});

test('duplicate expanded attributes refuse atomically even for same values or different prefixes', () => {
  const xml = '<r xmlns:p="u" xmlns:q="u"><t a="old"/></r>', s = XmlSnapshot.parse(xml), target = s.elements[1]!;
  for (const names of [['a', 'a'], ['p:new', 'q:new']]) {
    expect(() => s.setAttributes(names.map(name => ({ target, name, value: 'new' })))).toThrow(code('XML_ATTRIBUTE_DUPLICATE'));
    expect(s.setAttributes([])).toBe(xml);
  }
});

test('namespace declarations, unbound prefixes and malformed attribute QNames refuse', () => {
  const s = XmlSnapshot.parse('<r/>'), target = s.elements[0]!;
  for (const name of ['xmlns', 'xmlns:p', 'unbound:a', ':a', 'a:', 'a:b:c', '1name', 'a b', 'a="x"', '', 'a\u0000b']) {
    expect(() => s.setAttributes([{ target, name, value: 'v' }])).toThrow(code('XML_ATTRIBUTE_NAME'));
  }
  expect(s.setAttributes([{ target, name: '雪', value: 'v' }])).toBe('<r 雪="v"/>');
});

test('escaping preserves attribute values including both quotes, XML metacharacters and referenced whitespace', () => {
  const s = XmlSnapshot.parse('<r a=\'old\' b="old"/>'), target = s.elements[0]!;
  const value = '\'"<&>\t\r\n😀';
  const output = s.setAttributes([{ target, name: 'a', value }, { target, name: 'b', value }, { target, name: 'c', value }]);
  const root = parseXml(output).root;
  expect(root.attributes.a).toBe(value); expect(root.attributes.b).toBe(value); expect(root.attributes.c).toBe(value);
  expect(output).toContain("a='&#39;\"&lt;&amp;&gt;&#x9;&#xD;&#xA;😀'");
  expect(output).toContain('b="\'&quot;&lt;&amp;&gt;&#x9;&#xD;&#xA;😀"');
});

test('foreign, forged and stale targets refuse without changing snapshot or patch inputs', () => {
  const s = XmlSnapshot.parse(source), other = XmlSnapshot.parse(source), target = s.elements[1]!;
  for (const bad of [other.elements[1], { ...target }, null, undefined]) {
    expect(() => s.setAttributes([{ target: bad as typeof target, name: 'a', value: 'x' }])).toThrow(code('XML_ATTRIBUTE_TARGET'));
  }
  const patches = Object.freeze([Object.freeze({ target, name: 'a', value: 'next' })]);
  const output = s.setAttributes(patches); expect(patches[0]!.value).toBe('next');
  expect(() => XmlSnapshot.parse(output).setAttributes(patches)).toThrow(code('XML_ATTRIBUTE_TARGET'));
  expect(s.setAttributes([])).toBe(source);
});

test('invalid values and late invalid patches refuse with no partial output', () => {
  const s = XmlSnapshot.parse(source), target = s.elements[1]!;
  for (const value of ['\u0000', '\ud800', null, 123]) {
    let output: string | undefined;
    expect(() => { output = s.setAttributes([{ target, name: 'fresh', value: 'good' }, { target, name: 'a', value: value as string }]); }).toThrow(OoxmlError);
    expect(output).toBeUndefined(); expect(s.setAttributes([])).toBe(source);
  }
  expect(() => s.setAttributes(null as never)).toThrow(code('XML_ATTRIBUTE_PATCH'));
  expect(() => s.setAttributes([null] as never)).toThrow(code('XML_ATTRIBUTE_PATCH'));
});

test('root and descendant attribute edits preserve astral text, comments, PI, CRLF and attribute delimiters', () => {
  const xml = '<?xml version="1.0"?>\r\n<?pi keep?><r a="old>text">😀<!--keep--><c a=\'old\' /> tail\r\n</r>';
  const s = XmlSnapshot.parse(xml);
  expect(s.setAttributes([{ target: s.elements[1]!, name: 'a', value: '雪' }, { target: s.elements[0]!, name: 'a', value: 'new>text' }])).toBe(xml.replace('old>text', 'new&gt;text').replace("a='old'", "a='雪'"));
});

test('attribute patch input ignores caller iterators and refuses output or batch resource overflow', () => {
  const s = XmlSnapshot.parse('<r/>'), target = s.elements[0]!;
  const patches = [{ target, name: 'a', value: 'v' }];
  patches[Symbol.iterator] = function* () { throw new Error('do not iterate'); };
  expect(s.setAttributes(patches)).toBe('<r a="v"/>');
  expect(() => s.setAttributes([{ target, name: 'a', value: '&'.repeat(2 * 1024 * 1024) }])).toThrow(code('XML_ATTRIBUTE_LIMIT'));
  expect(() => s.setAttributes(new Array(100001).fill(patches[0]))).toThrow(code('XML_ATTRIBUTE_LIMIT'));
  expect(s.setAttributes([])).toBe('<r/>');
});

test('attribute span scanning uses XML whitespace, not Unicode JavaScript whitespace', () => {
  const name = '\ufeffname', xml = `<r ${name} = 'old' tail="keep"/>`, s = XmlSnapshot.parse(xml);
  expect(s.setAttributes([{ target: s.elements[0]!, name, value: 'new' }])).toBe(`<r ${name} = 'new' tail="keep"/>`);
});

test('many independent attribute edits retain order and avoid rescanning every attribute per patch', () => {
  const xml = '<r ' + Array.from({ length: 1500 }, (_, i) => `a${i}="old"`).join(' ') + '/>';
  const s = XmlSnapshot.parse(xml), target = s.elements[0]!;
  const output = s.setAttributes(Array.from({ length: 1500 }, (_, i) => ({ target, name: `a${i}`, value: `new${i}` })));
  expect(output).toBe('<r ' + Array.from({ length: 1500 }, (_, i) => `a${i}="new${i}"`).join(' ') + '/>');
});

test('four canonical attribute cases execute and incorrect output bytes fail predicates', async () => {
  const { fixturesRoot } = await import('../../scripts/fixture-inputs.ts'), { join } = await import('node:path');
  const { selectSharedScenarios, executeAcceptance } = await import('../../scripts/gherkin.ts');
  const { bindings } = await import('../acceptance/steps.ts'), { scenarioIds } = await import('../acceptance/xml-attributes.ts');
  const path = 'workflows/xml/editing.feature', text = await Bun.file(join(fixturesRoot(), path)).text();
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  const run = (source: string) => executeAcceptance({ root: '.', features: [selectSharedScenarios(path, source, scenarioIds)], counts: { features: count(1), scenarios: count(2), cases: count(4), steps: count(12) } }, bindings, 'xml-attributes-unit');
  const good = await run(text); expect(good.failures).toEqual([]); expect(good.counts.cases.passed).toBe(4); expect(good.counts.cases.planned).toBe(12);
  const bad = await run(text.replace('Then the complete output bytes equal <output>', 'Then the complete output bytes equal wrong'));
  expect(bad.counts.cases.failed).toBe(3); expect(bad.counts.steps.failed).toBe(3); expect(bad.counts.steps.undefined).toBe(0); expect(bad.counts.steps.ambiguous).toBe(0);
});

test('attribute duplicate refusal binding rejects fabricated success and untyped errors', async () => {
  const { fixturesRoot } = await import('../../scripts/fixture-inputs.ts'), { join } = await import('node:path');
  const { selectSharedScenarios, executeAcceptance } = await import('../../scripts/gherkin.ts');
  const { bindings } = await import('../acceptance/steps.ts');
  const path = 'workflows/xml/editing.feature', text = await Bun.file(join(fixturesRoot(), path)).text();
  const feature = selectSharedScenarios(path, text, ['@id-xml-go-attribute-batch-refusal']);
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  for (const success of [true, false]) {
    const corrupted = bindings.map(b => b.pattern.test('one Go lexical edit batch sets a of the first t element to x and to y') ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => {
      await b.run(c, ...captures); const s = c.state as { output?: string; error?: unknown };
      s.error = success ? undefined : new Error('untyped'); if (success) s.output = '<r/>';
    } } : b);
    const result = await executeAcceptance({ root: '.', features: [feature], counts: { features: count(1), scenarios: count(1), cases: count(1), steps: count(3) } }, corrupted, 'xml-attribute-refusal-control');
    expect(result.counts.cases.failed).toBe(1); expect(result.counts.steps.failed).toBe(1); expect(result.counts.steps.undefined).toBe(0); expect(result.counts.steps.ambiguous).toBe(0);
  }
});
