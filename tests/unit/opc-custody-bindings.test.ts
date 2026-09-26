import { expect, test } from 'bun:test';
import { join } from 'node:path';
import { fixturesRoot } from '../../scripts/fixture-inputs.ts';
import { executeAcceptance, selectSharedScenarios, type StepBinding } from '../../scripts/gherkin.ts';
import { bindings, cleanup } from '../acceptance/steps.ts';

const ids = ['@id-bun-opc-open-refusal', '@id-bun-opc-detached-byte-copies', '@id-bun-opc-preserve-utf16le-edit', '@id-bun-opc-async-transaction-refusal', '@id-bun-opc-thenable-transaction-result', '@id-bun-opc-save-invalid-target-custody', '@id-bun-opc-symlink-destination-refusal'];
async function run(change: (s: string) => string = s => s, activeBindings: StepBinding[] = bindings) {
  const path = 'workflows/package/preservation.feature';
  const f = selectSharedScenarios(path, change(await Bun.file(join(fixturesRoot(), path)).text()), ids);
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  try { return await executeAcceptance({ root: '.', features: [f], counts: { features: count(1), scenarios: count(7), cases: count(11), steps: count(107) } }, activeBindings, 'opc-custody-unit'); }
  finally { await cleanup(); }
}
function failedPredicate(r: Awaited<ReturnType<typeof run>>) {
  expect(r.counts.cases.failed).toBeGreaterThan(0); expect(r.counts.steps.failed).toBeGreaterThan(0);
  expect(r.counts.steps.undefined).toBe(0); expect(r.counts.steps.ambiguous).toBe(0);
}
test('eleven canonical OPC custody cases execute, including backgrounds and save-path outcomes', async () => {
  const r = await run(); expect(r.failures).toEqual([]); expect(r.counts.cases.passed).toBe(11); expect(r.counts.cases.planned).toBe(3);
});
test('changed OPC error-code, message, content and BOM predicates fail', async () => {
  for (const [from, to] of [
    ['code opc-async-transaction', 'code opc-wrong-code'],
    ['message contains synchronous edits', 'message contains wrong message'],
    ['contains the UTF-8 text Alpha', 'contains the UTF-8 text Wrong'],
    ['starts with hexadecimal bytes FF FE', 'starts with hexadecimal bytes FE FF'],
  ]) failedPredicate(await run(s => { expect(s.includes(from!)).toBe(true); return s.replace(from!, to!); }));
});

test('OPC refusal archives contain the exact background and named mutation', async () => {
  const { readZip } = await import('../../src/opc/zip.ts');
  const { parseXml } = await import('../../src/xml/index.ts');
  const snapshots: Uint8Array[] = [];
  const checking = bindings.map(b => b.pattern.test('Bun OpcPackage opens its archive bytes') ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => {
    await b.run(c, ...captures);
    const s = c.state as { archive: Uint8Array }; snapshots.push(s.archive.slice());
  } } : b);
  const result = await run(s => s, checking); expect(result.failures).toEqual([]); expect(snapshots).toHaveLength(5);
  for (let i = 0; i < snapshots.length; i++) {
    const parts = readZip(snapshots[i]!);
    const document = i === 0 ? 'word/%66oo.xml' : 'word/document.xml';
    expect([...parts.keys()]).toEqual(['[Content_Types].xml', '_rels/.rels', document]);
    const decode = (name: string) => new TextDecoder('utf-8', { fatal: true }).decode(parts.get(name)!);
    expect(decode(document)).toBe('<?xml version="1.0" encoding="UTF-8"?><document>Alpha</document>');
    const ct = parseXml(decode('[Content_Types].xml')).root;
    expect(ct.namespaceURI).toBe('http://schemas.openxmlformats.org/package/2006/content-types');
    expect(ct.children.map(n => [n.localName, { ...n.attributes }])).toEqual([
      ['Default', { Extension: 'rels', ContentType: 'application/vnd.openxmlformats-package.relationships+xml' }],
      ['Default', { Extension: i === 2 ? 'rels' : 'xml', ContentType: 'application/xml' }],
      ['Override', { PartName: '/' + document, ContentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' }],
    ]);
    const rels = parseXml(decode('_rels/.rels')).root;
    expect(rels.namespaceURI).toBe('http://schemas.openxmlformats.org/package/2006/relationships');
    expect(rels.children.map(n => [n.localName, { ...n.attributes }])).toEqual(Array.from({ length: i === 4 ? 2 : 1 }, () => ['Relationship', {
      Id: 'rId1', Type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      Target: i < 2 ? 'word/%66oo.xml' : i === 3 ? 'word/missing.xml' : 'word/document.xml',
    }]));
  }
});

test('OPC predicates reject changed package bytes, callback flags, thenable identity and saved members', async () => {
  type State = import('../acceptance/opc-custody.ts').CustodyState;
  const mutations: [string, (s: State) => void][] = [
    ['Bun OpcPackage opens its archive bytes', s => { s.error = undefined; s.output = {}; }],
    ['Bun OpcPackage opens its archive bytes', s => { s.error = new Error('untyped'); }],
    ["every byte in the array returned by get for word/document.xml is overwritten with zero", s => { s.pkg!.set('word/document.xml', '<document>Wrong</document>'); }],
    ["every byte in the array returned by get for word/document.xml is overwritten with zero", s => { s.original![0] = s.original![0]! ^ 1; }],
    ['its serialized archive is read through the ZIP layer', s => { s.member![0] = 0; }],
    ['its serialized archive is read through the ZIP layer', s => { s.member = new Uint8Array([255, 254, 65, 0]); }],
    ['its transaction is called with an async callback that would replace Alpha with Beta and set a ran flag', s => { s.ran = true; }],
    ['a synchronous transaction replaces Alpha with Beta and returns that thenable object', s => { s.transactionResult = {}; }],
    ['a synchronous transaction replaces Alpha with Beta and returns that thenable object', s => { s.thenCalls = 1; }],
    ['a synchronous transaction replaces Alpha with Beta and returns that thenable object', s => { s.pkg!.set('word/document.xml', '<document>Alpha</document>'); }],
  ];
  for (const [step, mutate] of mutations) {
    const corrupt = bindings.map(b => b.pattern.test(step) ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => {
      await b.run(c, ...captures); mutate(c.state as State);
    } } : b);
    failedPredicate(await run(s => s, corrupt));
  }
});

test('OPC destination assertions detect overwritten files and replaced symlinks', async () => {
  const { unlink } = await import('node:fs/promises');
  type State = import('../acceptance/opc-custody.ts').CustodyState;
  for (const [step, mutate] of [
    ['the package is saved to the existing destination', async (s: State) => { await Bun.write(s.destination!, 'changed'); }],
    ['the unchanged package is saved through the symlink path', async (s: State) => { await Bun.write(s.destination!, 'changed'); }],
    ['the unchanged package is saved through the symlink path', async (s: State) => { await unlink(s.link!); await Bun.write(s.link!, s.original!); }],
  ] as const) {
    const corrupt = bindings.map(b => b.pattern.test(step) ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => {
      await b.run(c, ...captures); const s = c.state as State;
      // The save binding covers both destinations; restrict this mutation to the requested one.
      if (Boolean(s.link) === step.includes('symlink')) await mutate(s);
    } } : b);
    failedPredicate(await run(s => s, corrupt));
  }
});

test('OPC temporary destinations are cleaned after success and predicate failure', async () => {
  const { lstat } = await import('node:fs/promises');
  const paths: string[] = [];
  const checking = bindings.map(b => b.pattern.test('an existing destination file contains the base package archive bytes') ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => {
    await b.run(c, ...captures); paths.push((c.state as { destination: string }).destination);
  } } : b);
  expect((await run(s => s, checking)).failures).toEqual([]);
  failedPredicate(await run(s => s.replace('message contains Missing target', 'message contains wrong'), checking));
  expect(paths).toHaveLength(4);
  for (const path of paths) {
    await expect(lstat(path)).rejects.toMatchObject({ code: 'ENOENT' });
    await expect(lstat(join(path, '..'))).rejects.toMatchObject({ code: 'ENOENT' });
  }
});
