import { expect, test } from 'bun:test';
import { join } from 'node:path';
import { fixturesRoot } from '../../scripts/fixture-inputs.ts';
import { executeAcceptance, selectSharedScenarios, type StepBinding } from '../../scripts/gherkin.ts';
import { bindings } from '../acceptance/steps.ts';

const generalized=(await Bun.file(join(fixturesRoot(),'workflows/package/zip32.feature')).text()).includes('@profile-zip32-refusal-reasons');
const ids = ['@id-zip-crc32-standard-vector', '@id-bun-zip32-reader-refusal', '@id-bun-zip32-writer-refusal', '@id-bun-zip32-configured-bounds'];
async function run(change: (s: string) => string = s => s, activeBindings: StepBinding[] = bindings) {
  const path = 'workflows/package/zip32.feature';
  const feature = selectSharedScenarios(path, change(await Bun.file(join(fixturesRoot(), path)).text()), ids);
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  return executeAcceptance({ root: '.', features: [feature], counts: {
    features: count(1), scenarios: count(4), cases: count(20), steps: count(91),
  } }, activeBindings, 'zip32-unit');
}

test('twenty canonical ZIP32 checksum, refusal and configured-bound cases execute', async () => {
  const result = await run();
  expect(result.failures).toEqual([]);
  expect(result.counts.cases.passed).toBe(20);
  expect(result.counts.cases.planned).toBe(4);
  expect(result.counts.steps.undefined).toBe(0);
});

test('wrong CRC, error code and message expectations fail bound ZIP32 predicates', async () => {
  for (const [from, to] of [
    ['hexadecimal CBF43926', 'hexadecimal DEADBEEF'],
    ['zip-crc-mismatch', 'zip-wrong-code'],
    generalized?['reason <code>','reason zip-wrong-reason']:['| CRC                      |', '| wrong message            |'],
  ]) {
    const result = await run(s => { expect(s.includes(from!)).toBe(true); return s.replaceAll(from!, to!); });
    expect(result.counts.cases.failed).toBeGreaterThan(0);
    expect(result.counts.steps.failed).toBeGreaterThan(0);
    expect(result.counts.steps.undefined).toBe(0);
    expect(result.counts.steps.ambiguous).toBe(0);
  }
});

test('ZIP32 refusal fixtures encode the specified local, central and end-record mutations', async () => {
  const { readerSample } = await import('../acceptance/zip32.ts');
  const { inflateRawSync } = await import('node:zlib');
  const { crc32Test } = await import('../fixtures/zip32.ts');
  const encode = (s: string) => new TextEncoder().encode(s), decode = (b: Uint8Array) => new TextDecoder().decode(b);
  const sample = (mutation: string, value = 'payload') => {
    const bytes = readerSample([{ name: 'word/document.xml', blob: encode(value) }], mutation);
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const end = bytes.length - 22, central = view.getUint32(end + 16, true);
    const data = 30 + view.getUint16(26, true) + view.getUint16(28, true);
    expect(view.getUint32(0, true)).toBe(0x04034b50);
    expect(view.getUint32(central, true)).toBe(0x02014b50);
    expect(view.getUint32(end, true)).toBe(0x06054b50);
    return { bytes, view, end, central, data };
  };
  const plain = sample('none');
  expect(plain.view.getUint16(6, true)).toBe(0x800);
  expect(plain.view.getUint16(8, true)).toBe(8);
  expect(plain.view.getUint32(14, true)).toBe(crc32Test(encode('payload')));
  expect(decode(inflateRawSync(plain.bytes.subarray(plain.data, plain.central)))).toBe('payload');
  const encrypted = sample('general-purpose bit 0 is set in both headers');
  expect(encrypted.view.getUint16(6, true)).toBe(0x801);
  expect(encrypted.view.getUint16(encrypted.central + 8, true)).toBe(0x801);
  const unsupported = sample('both methods are 12 and payload bytes are stored uncompressed');
  expect(unsupported.view.getUint16(8, true)).toBe(12);
  expect(unsupported.view.getUint16(unsupported.central + 10, true)).toBe(12);
  expect(decode(unsupported.bytes.subarray(unsupported.data, unsupported.central))).toBe('payload');
  const disk = sample('end record disk number is 1');
  expect(disk.view.getUint16(disk.end + 4, true)).toBe(1);
  const counts = sample('both end-record counts are 65535 without ZIP64 records');
  expect(counts.view.getUint16(counts.end + 8, true)).toBe(65535);
  expect(counts.view.getUint16(counts.end + 10, true)).toBe(65535);
  expect(counts.central + counts.view.getUint32(counts.end + 12, true)).toBe(counts.end);
  const mismatch = sample('local name is word/other.xml');
  expect(decode(mismatch.bytes.subarray(30, mismatch.data))).toBe('word/other.xml');
  expect(decode(mismatch.bytes.subarray(mismatch.central + 46, mismatch.central + 46 + mismatch.view.getUint16(mismatch.central + 28, true)))).toBe('word/document.xml');
  const crc = sample('both CRC fields are hexadecimal DEADBEEF');
  expect(crc.view.getUint32(14, true)).toBe(0xdeadbeef);
  expect(crc.view.getUint32(crc.central + 16, true)).toBe(0xdeadbeef);
  const stored = sample('method is STORED and all size fields are 99');
  expect(stored.view.getUint16(8, true)).toBe(0);
  expect(stored.view.getUint16(stored.central + 10, true)).toBe(0);
  for (const offset of [18, 22, stored.central + 20, stored.central + 24]) expect(stored.view.getUint32(offset, true)).toBe(99);
  expect(decode(stored.bytes.subarray(stored.data, stored.central))).toBe('payload');
  const overrun = sample('payload repeats A 4096 times but both expanded sizes are 32', 'A');
  expect(overrun.view.getUint32(22, true)).toBe(32);
  expect(overrun.view.getUint32(overrun.central + 24, true)).toBe(32);
  expect(overrun.view.getUint32(18, true)).toBe(overrun.central - overrun.data);
  expect(overrun.view.getUint32(overrun.central + 20, true)).toBe(overrun.central - overrun.data);
  expect(decode(inflateRawSync(overrun.bytes.subarray(overrun.data, overrun.central)))).toBe('A'.repeat(4096));
  expect(overrun.view.getUint32(14, true)).toBe(crc32Test(encode('A'.repeat(4096))));
  expect(overrun.view.getUint32(overrun.central + 16, true)).toBe(crc32Test(encode('A'.repeat(4096))));
  const tail = readerSample([{ name: 'word/document.xml', blob: encode('payload') }], 'a newline byte follows the complete uncommented archive');
  expect(tail.subarray(0, -1)).toEqual(plain.bytes); expect(tail.at(-1)).toBe(10);
});

test('ordered ZIP32 fixture pairs keep duplicates, case differences and noncanonical names', async () => {
  const { readerSample } = await import('../acceptance/zip32.ts');
  for (const names of [['word/document.xml', 'word/document.xml'], ['word/document.xml', 'WORD/document.xml'], ['../word/document.xml']]) {
    const specs = names.map((name, i) => ({ name, blob: new TextEncoder().encode(String(i)) }));
    const bytes = readerSample(specs, 'none'), view = new DataView(bytes.buffer);
    let cursor = view.getUint32(bytes.length - 6, true);
    const actual: string[] = [];
    for (let i = 0; i < names.length; i++) {
      const size = view.getUint16(cursor + 28, true);
      actual.push(new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + size)));
      cursor += 46 + size;
    }
    expect(actual).toEqual(names);
    expect(specs.map(s => s.name)).toEqual(names);
    expect(specs.map(s => new TextDecoder().decode(s.blob))).toEqual(names.map((_, i) => String(i)));
  }
});

test('ZIP32 predicates reject fabricated success, untyped errors and caller-buffer mutation', async () => {
  type State = { error?: unknown; output?: unknown; archive?: Uint8Array };
  const steps = ['the ZIP reader reads the sample with default limits', 'the ZIP writer writes the entries with default options', 'the ZIP reader reads the archive with only maxEntries set to 1'];
  for (const step of steps) for (const mutate of [
    (s: State) => { s.error = undefined; s.output = new Map(); },
    (s: State) => { s.error = new Error('untyped'); },
    ...(step.includes('the ZIP reader') ? [(s: State) => { if (s.archive) s.archive[0] = s.archive[0]! ^ 1; }] : []),
  ]) {
    const corrupted = bindings.map(b => b.pattern.test(step) ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => {
      await b.run(c, ...captures); mutate(c.state as State);
    } } : b);
    const result = await run(s => s, corrupted);
    expect(result.counts.cases.failed).toBeGreaterThan(0);
    expect(result.counts.steps.failed).toBeGreaterThan(0);
    expect(result.counts.steps.undefined).toBe(0);
    expect(result.counts.steps.ambiguous).toBe(0);
  }
});
