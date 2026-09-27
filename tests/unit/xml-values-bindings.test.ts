import { expect, test } from 'bun:test';
import { join } from 'node:path';
import { fixturesRoot } from '../../scripts/fixture-inputs.ts';
import { executeAcceptance, selectSharedScenarios, type StepBinding } from '../../scripts/gherkin.ts';
import type { XmlDocument } from '../../src/xml/index.ts';
import { bindings } from '../acceptance/steps.ts';

const ids = [
  '@id-xml-entity-values', '@id-xml-stylesheet-processing-instruction',
  '@id-xml-expanded-attribute-lookup', '@id-xml-implicit-xml-prefix',
  '@id-xml-prototype-safe-attributes', '@id-xml-immutable-namespace-metadata',
  '@id-xml-escaping-values', '@id-xml-escaping-invalid-character',
  '@id-xml-escaping-whitespace-roundtrip', '@id-xml-typed-parse-error',
];
function assertFailedPredicates(result: Awaited<ReturnType<typeof executeAcceptance>>) {
  expect(result.counts.cases.failed).toBeGreaterThan(0);
  expect(result.failures.length).toBeGreaterThan(0);
  expect(result.counts.steps.failed).toBeGreaterThan(0);
  expect(result.counts.steps.undefined).toBe(0);
  expect(result.counts.steps.ambiguous).toBe(0);
}
async function run(sourceChange: (source: string) => string = s => s, activeBindings: StepBinding[] = bindings) {
  const path = 'workflows/xml/parsing.feature';
  const feature = selectSharedScenarios(path, sourceChange(await Bun.file(join(fixturesRoot(), path)).text()), ids);
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  return executeAcceptance({ root: '.', features: [feature], counts: {
    features: count(1), scenarios: count(10), cases: count(11), steps: count(39),
  } }, activeBindings, 'xml-values-unit');
}

test('eleven XML value cases execute with every Then asserted and other operations left planned', async () => {
  const result = await run();
  expect(result.failures).toEqual([]);
  expect(result.counts.cases.passed).toBe(11);
  expect(result.counts.cases.planned).toBe(4);
  const executed = result.features.flatMap(f => f.scenarios.flatMap(s => s.cases)).filter(c => c.lifecycle === 'implemented');
  expect(executed).toHaveLength(11);
  expect(executed.every(c => c.steps.every(s => s.status === 'passed'))).toBe(true);
});

test('changed expected XML values and namespace-table cells fail instead of passing vacuously', async () => {
  const changes: [string, string][] = [
    ['root attribute a equals JSON "\\\"\'"', 'root attribute a equals JSON "wrong"'],
    ['root text equals JSON "AA&<>"', 'root text equals JSON "wrong"'],
    ['root qualified name equals r', 'root qualified name equals wrong'],
    ['root namespace URI equals JSON ""', 'root namespace URI equals JSON "wrong"'],
    ['root attribute xml:lang equals JSON "en"', 'root attribute xml:lang equals JSON "wrong"'],
    ['implicit xml namespace URI is http://www.w3.org/XML/1998/namespace', 'implicit xml namespace URI is urn:wrong'],
    ['namespace recorded for a:id is urn:a', 'namespace recorded for a:id is urn:wrong'],
    ['| root    | id    | urn:default                                    | null', '| root    | id    | urn:default                                    | "wrong"'],
    ['escaped string equals JSON <escaped_json>', 'escaped string equals JSON "wrong"'],
    ['decoded text and attribute both equal JSON "x\\r\\n\\ty"', 'decoded text and attribute both equal JSON "wrong"'],
  ];
  for (const [before, after] of changes) {
    const result = await run(source => { expect(source.includes(before)).toBe(true); return source.replace(before, after); });
    assertFailedPredicates(result);
  }
});

test('XML value assertions reject corrupted parser, metadata and escaping results', async () => {
  type State = { document?: XmlDocument; error?: unknown; escaped?: string };
  const mutations: [string, (s: State) => void][] = [
    ['the XML values input is parsed', s => { if (s.document) s.document.root.attributes = { ...s.document.root.attributes }; }],
    ['the XML values input is parsed', s => {
      if (s.document) Object.setPrototypeOf(s.document.root.attributes, { constructor: 'safe' });
      if (s.document) Reflect.deleteProperty(s.document.root.attributes, 'constructor');
    }],
    ['the XML values input is parsed', s => {
      if (s.document) s.document.root.attributeNamespaces = Object.assign(Object.create(null), s.document.root.attributeNamespaces);
    }],
    ['the XML values input is parsed', s => {
      if (s.document) s.document.root.attributeNamespaces = Object.freeze({ ...s.document.root.attributeNamespaces });
    }],
    ['the XML values input is parsed', s => { if (s.error) s.error = new Error('untyped'); }],
    ['the value is escaped for XML text content', s => { if (s.error) { s.error = undefined; s.escaped = ''; } }],
  ];
  for (const [step, mutate] of mutations) {
    const corrupted = bindings.map(binding => binding.pattern.test(step) ? {
      ...binding, run: async (context: Record<string, unknown>, ...captures: string[]) => {
        await binding.run(context, ...captures); mutate(context.state as State);
      },
    } : binding);
    const result = await run(s => s, corrupted);
    assertFailedPredicates(result);
  }
});
