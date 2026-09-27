import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { crc32, readZip, writeZip, type ZipLimits } from '../../src/opc/zip.ts';
import { buildZip, type ZipMemberSpec } from '../fixtures/zip32.ts';

const encoder = new TextEncoder();
type State = { entries?: ZipMemberSpec[]; archive?: Uint8Array; before?: Uint8Array; output?: unknown; error?: unknown; checksum?: number; checksumInput?: Uint8Array };
const state = (c: Record<string, unknown>) => c.state as State;
export const scenarioIds = ['@id-zip-crc32-standard-vector', '@id-bun-zip32-reader-refusal', '@id-bun-zip32-writer-refusal', '@id-bun-zip32-configured-bounds'];
export function readerSample(entries: ZipMemberSpec[], mutation: string): Uint8Array {
  const specs = entries.map(e => ({ ...e, blob: e.blob?.slice() }));
  assert(specs.length > 0);
  const first = specs[0]!;
  let options = {};
  switch (mutation) {
    case 'none': break;
    case 'general-purpose bit 0 is set in both headers': first.flags = 0x0801; break;
    case 'both methods are 12 and payload bytes are stored uncompressed': first.method = 12; break;
    case 'end record disk number is 1': options = { diskNumber: 1 }; break;
    case 'both end-record counts are 65535 without ZIP64 records': options = { diskEntries: 65535, totalEntries: 65535 }; break;
    case 'local name is word/other.xml': first.localName = 'word/other.xml'; break;
    case 'both CRC fields are hexadecimal DEADBEEF': first.crc = 0xdeadbeef; break;
    case 'method is STORED and all size fields are 99': Object.assign(first, { method: 0, fileSize: 99, compressedSize: 99, localFileSize: 99, localCompressedSize: 99 }); break;
    case 'payload repeats A 4096 times but both expanded sizes are 32': Object.assign(first, { blob: encoder.encode('A'.repeat(4096)), fileSize: 32, localFileSize: 32 }); break;
    case 'a newline byte follows the complete uncommented archive': return Uint8Array.from([...buildZip(specs), 10]);
    default: throw new Error('Unknown ZIP32 sample mutation: ' + mutation);
  }
  return buildZip(specs, options);
}
function entries(json: string | undefined): ZipMemberSpec[] {
  assert(json); const rows: unknown = JSON.parse(json); assert(Array.isArray(rows) && rows.length > 0);
  return rows.map(row => { assert(Array.isArray(row) && row.length === 2 && row.every(x => typeof x === 'string')); return { name: row[0]!, blob: encoder.encode(row[1]!) }; });
}
function read(s: State, limits: ZipLimits = {}) {
  assert(s.archive); s.before = s.archive.slice();
  try { s.output = readZip(s.archive, limits); } catch (error) { s.error = error; }
}
export const bindings: StepBinding[] = [
  { pattern: /^checksum input is the UTF-8 string (".+")$/, run: (c, json) => { state(c).checksumInput = encoder.encode(JSON.parse(json!)); } },
  { pattern: /^its ZIP CRC32 is calculated$/, run: c => { const s = state(c); assert(s.checksumInput); s.checksum = crc32(s.checksumInput); } },
  { pattern: /^the unsigned checksum equals hexadecimal ([A-Fa-f0-9]+)$/, run: (c, hex) => { assert.equal(state(c).checksum, Number.parseInt(hex!, 16)); } },
  { pattern: /^a ZIP32 reader sample with these ordered member and payload pairs encoded as JSON (.+)$/, run: (c, json) => { state(c).entries = entries(json); } },
  { pattern: /^the sample has the mutation (.+)$/, run: (c, mutation) => { const s = state(c); assert(s.entries); s.archive = readerSample(s.entries, mutation!); } },
  { pattern: /^the ZIP reader reads the sample with default limits$/, run: c => { read(state(c)); } },
  { pattern: /^ordered writer entries are encoded as JSON (.+)$/, run: (c, json) => { state(c).entries = entries(json); } },
  { pattern: /^the ZIP writer writes the entries with default options$/, run: c => {
    const s = state(c); assert(s.entries);
    const input = new Map(s.entries.map(e => [e.name, e.blob!])), before = [...input].map(([n, b]) => [n, b.slice()]);
    try { s.output = writeZip(input); } catch (error) { s.error = error; }
    assert.deepEqual([...input], before);
  } },
  { pattern: /^a raw-DEFLATE ZIP32 archive contains a.bin with 4096 A bytes followed by b.bin with two b bytes$/, run: c => {
    const s = state(c); s.archive = buildZip([{ name: 'a.bin', blob: encoder.encode('A'.repeat(4096)) }, { name: 'b.bin', blob: encoder.encode('bb') }]);
    const control = readZip(s.archive); assert.deepEqual([...control.keys()], ['a.bin', 'b.bin']);
    assert.deepEqual(control.get('a.bin'), encoder.encode('A'.repeat(4096))); assert.deepEqual(control.get('b.bin'), encoder.encode('bb'));
  } },
  { pattern: /^the ZIP reader reads the archive with only (maxArchiveBytes|maxEntries|maxEntryBytes|maxTotalBytes|maxCompressionRatio) set to (.+)$/, run: (c, limit, value) => {
    const s = state(c); assert(s.archive);
    const n = value === 'archive byte length minus 1' ? s.archive.length - 1 : Number(value);
    assert(Number.isSafeInteger(n) && n > 0); read(s, { [limit!]: n });
  } },
];
