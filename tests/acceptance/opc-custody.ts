import assert from 'node:assert/strict';
import { mkdtemp, rm, symlink, lstat, readlink } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import type { AcceptanceStep, StepBinding } from '../../scripts/gherkin.ts';
import { OpcPackage } from '../../src/opc/package.ts';
import { readZip, writeZip } from '../../src/opc/zip.ts';
import { escapeAttribute } from '../../src/xml/index.ts';
import { utf16, utf8 } from '../fixtures/admission.ts';

const doc = 'word/document.xml';
const roots = new Set<string>();
export async function cleanup() { const paths = [...roots]; await Promise.all(paths.map(async path => { await rm(path, { recursive: true, force: true }); roots.delete(path); })); }
type ContentType = { kind: string; key: string; type: string };
export type CustodyState = {
  part?: string; payload?: Uint8Array; contentNamespace?: string; types?: ContentType[]; relationNamespace?: string;
  relation?: { id: string; type: string; target: string }; duplicateRelation?: boolean;
  archive?: Uint8Array; before?: Uint8Array; original?: Uint8Array; pkg?: OpcPackage; output?: unknown; error?: unknown;
  ran?: boolean; thenCalls?: number; thenable?: { then: () => never }; transactionResult?: unknown;
  saved?: Uint8Array; member?: Uint8Array; destination?: string; link?: string;
};
const state = (c: Record<string, unknown>) => c.state as CustodyState;
function pkg(s: CustodyState) { assert(s.pkg); return s.pkg; }
function archive(s: CustodyState) {
  assert(s.part && s.payload && s.contentNamespace && s.types && s.relationNamespace && s.relation);
  const a = escapeAttribute, r = s.relation;
  const relation = `<Relationship Id="${a(r.id)}" Type="${a(r.type)}" Target="${a(r.target)}"/>`;
  const types = s.types.map(t => `<${t.kind} ${t.kind === 'Default' ? 'Extension' : 'PartName'}="${a(t.key)}" ContentType="${a(t.type)}"/>`).join('');
  return writeZip(new Map([
    ['[Content_Types].xml', utf8(`<?xml version="1.0" encoding="UTF-8"?><Types xmlns="${a(s.contentNamespace)}">${types}</Types>`)],
    ['_rels/.rels', utf8(`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${a(s.relationNamespace)}">${relation}${s.duplicateRelation ? relation : ''}</Relationships>`)],
    [s.part, s.payload],
  ]));
}
async function open(s: CustodyState) { s.archive ??= archive(s); s.original ??= s.archive.slice(); s.pkg = await OpcPackage.open(s.archive); }
function retainSource(s: CustodyState) { s.before = s.archive!.slice(); }
async function destination(s: CustodyState) {
  const root = await mkdtemp(join(tmpdir(), 'bun-opc-custody-')); roots.add(root);
  s.destination = join(root, 'target.docx'); s.archive = archive(s); s.original = s.archive.slice();
  await Bun.write(s.destination, s.archive);
}
function savedCustody(s: CustodyState) {
  assert(s.original && s.saved);
  const before = readZip(s.original), after = readZip(s.saved);
  assert.deepEqual([...after.keys()], [...before.keys()]);
  for (const [name, bytes] of before) if (name !== doc) assert.deepEqual(after.get(name), bytes);
}
export const scenarioIds = ['@id-bun-opc-open-refusal', '@id-bun-opc-detached-byte-copies', '@id-bun-opc-preserve-utf16le-edit', '@id-bun-opc-async-transaction-refusal', '@id-bun-opc-thenable-transaction-result', '@id-bun-opc-save-invalid-target-custody', '@id-bun-opc-symlink-destination-refusal'];
export const bindings: StepBinding[] = [
  { pattern: /^a ZIP contains word\/document.xml with UTF-8 XML text (.+)$/, run: (c, text) => { Object.assign(state(c), { part: doc, payload: utf8(text!) }); } },
  { pattern: /^its content types use namespace (\S+)$/, run: (c, ns) => { state(c).contentNamespace = ns; } },
  { pattern: /^its content types have these defaults and overrides$/, run: c => {
    const rows = (c.step as AcceptanceStep).argument?.dataTable; assert(rows && rows.length > 1);
    assert.deepEqual(rows[0], ['kind', 'key', 'content_type']);
    state(c).types = rows.slice(1).map(row => { assert.equal(row.length, 3); const [kind, key, type] = row; assert(kind === 'Default' || kind === 'Override'); assert(key && type); return { kind, key, type }; });
  } },
  { pattern: /^_rels\/\.rels uses namespace (\S+)$/, run: (c, ns) => { state(c).relationNamespace = ns; } },
  { pattern: /^its root relationship is (\S+) of type (\S+) targeting (\S+)$/, run: (c, id, type, target) => { assert(id && type && target); state(c).relation = { id, type, target }; } },
  { pattern: /^the base package has the mutation (.+)$/, run: (c, mutation) => {
    const s = state(c); assert(s.relation && s.types);
    switch (mutation) {
      case 'rename the document, content-type override and relationship target to word/%66oo.xml':
        s.part = 'word/%66oo.xml'; s.relation.target = s.part;
        for (const t of s.types) if (t.kind === 'Override' && t.key === '/' + doc) t.key = '/' + s.part; break;
      case 'change only the root relationship target to word/%66oo.xml': s.relation.target = 'word/%66oo.xml'; break;
      case 'replace the xml default with a second rels default of type application/xml': {
        const t = s.types.find(t => t.kind === 'Default' && t.key === 'xml'); assert(t); t.key = 'rels'; t.type = 'application/xml'; break;
      }
      case 'change only the root relationship target to word/missing.xml': s.relation.target = 'word/missing.xml'; break;
      case 'duplicate the complete rId1 root relationship': s.duplicateRelation = true; break;
      default: throw new Error('Unknown OPC sample mutation: ' + mutation);
    }
  } },
  { pattern: /^the package editor opens its archive bytes$/, run: async c => {
    const s = state(c); s.archive = archive(s); retainSource(s);
    try { s.output = await OpcPackage.open(s.archive); } catch (error) { s.error = error; }
  } },
  { pattern: /^the package editor (?:opens|has opened) the base archive bytes$/, run: async c => { await open(state(c)); } },
  { pattern: /^every byte in the caller's original archive array is overwritten with zero$/, run: c => { const s = state(c); assert(s.archive); s.archive.fill(0); assert(s.archive.every(b => b === 0)); } },
  { pattern: /^every byte in the array returned by get for word\/document.xml is overwritten with zero$/, run: c => { const b = pkg(state(c)).get(doc); assert(b && b.length > 0); b.fill(0); assert(b.every(v => v === 0)); } },
  { pattern: /^a fresh get of word\/document.xml contains the UTF-8 text (.+)$/, run: (c, text) => { const b = pkg(state(c)).get(doc); assert(b); assert(new TextDecoder('utf-8', { fatal: true }).decode(b).includes(text!)); } },
  { pattern: /^serializing the package returns the exact original archive bytes$/, run: async c => {
    const s = state(c); assert(s.original); s.saved = pkg(s).toBytes(); assert.deepEqual(s.saved, s.original);
    assert.deepEqual((await OpcPackage.open(s.saved)).toBytes(), s.original);
  } },
  { pattern: /^word\/document.xml instead has a little-endian BOM and UTF-16LE text (.+)$/, run: (c, text) => { state(c).payload = utf16(text!); } },
  { pattern: /^the package editor opens the package and sets word\/document.xml to its decoded text with Alpha replaced by Beta$/, run: async c => { const s = state(c); await open(s); const text = pkg(s).text(doc); assert(text.includes('Alpha')); pkg(s).set(doc, text.replace('Alpha', 'Beta')); } },
  { pattern: /^its serialized archive is read through the ZIP layer$/, run: c => { const s = state(c); s.saved = pkg(s).toBytes(); s.member = readZip(s.saved).get(doc); assert(s.member); savedCustody(s); } },
  { pattern: /^the saved document member starts with hexadecimal bytes ([A-Fa-f0-9 ]+)$/, run: (c, bytes) => { const s = state(c); assert(s.member); const expected = bytes!.split(' ').map(b => parseInt(b, 16)); assert.deepEqual([...s.member.slice(0, expected.length)], expected); } },
  { pattern: /^decoding the member as UTF-16LE contains Beta and encoding="UTF-16"$/, run: async c => {
    const s = state(c); assert(s.member && s.saved); const text = new TextDecoder('utf-16le', { fatal: true }).decode(s.member);
    assert(text.includes('Beta')); assert(text.includes('encoding="UTF-16"')); assert(!text.includes('Alpha'));
    assert.equal((await OpcPackage.open(s.saved)).text(doc), text);
  } },
  {pattern:/^a deferred transaction is requested with a callback that would replace Alpha with Beta and set a ran flag$/,run:c=>{
    const s=state(c);s.ran=false;retainSource(s);try{s.output=pkg(s).transactionWithMode('deferred',()=>{s.ran=true;pkg(s).set(doc,pkg(s).text(doc).replace('Alpha','Beta'));});}catch(error){s.error=error;}
  }},
  {pattern:/^the current package archive and caller source bytes remain unchanged$/,run:c=>{const s=state(c);assert(s.archive&&s.before&&s.original);assert.deepEqual(s.archive,s.before);assert.deepEqual(pkg(s).toBytes(),s.original);}},
  {pattern:/^an opaque token has an evaluation hook that throws if invoked$/,run:c=>{const s=state(c);s.thenCalls=0;s.thenable={then:()=>{s.thenCalls!++;throw Error('must not be evaluated');}};}},
  {pattern:/^an immediate transaction replaces Alpha with Beta and returns that token$/,run:c=>{const s=state(c);assert(s.thenable);s.transactionResult=pkg(s).transactionWithMode('immediate',()=>{pkg(s).set(doc,pkg(s).text(doc).replace('Alpha','Beta'));return s.thenable;});}},
  {pattern:/^the returned token has the original identity and its evaluation count is (zero|one)$/,run:(c,count)=>{const s=state(c);assert(s.thenable);assert.equal(s.transactionResult,s.thenable);assert.equal(s.thenCalls,count==='zero'?0:1);}},
  {pattern:/^saving and reopening reads Beta with every unrelated member payload unchanged$/,run:async c=>{
    const s=state(c);assert(s.original);const root=await mkdtemp(join(tmpdir(),'portable-transaction-'));roots.add(root);const file=join(root,'saved.docx');await pkg(s).save(file);s.saved=Uint8Array.from(await Bun.file(file).bytes());savedCustody(s);assert((await OpcPackage.open(s.saved)).text(doc).includes('Beta'));assert.equal(s.thenCalls,0);assert.deepEqual(s.archive,s.original);
  }},
  { pattern: /^its transaction is called with an async callback that would replace Alpha with Beta and set a ran flag$/, run: c => {
    const s = state(c); s.ran = false; retainSource(s);
    try { s.output = pkg(s).transaction(async () => { s.ran = true; pkg(s).set(doc, pkg(s).text(doc).replace('Alpha', 'Beta')); }); }
    catch (error) { s.error = error; }
  } },
  { pattern: /^the ran flag is false and the document text still contains Alpha$/, run: c => { const s = state(c); assert.equal(s.ran, false); assert(pkg(s).text(doc).includes('Alpha')); assert.deepEqual(pkg(s).toBytes(), s.original); } },
  { pattern: /^a thenable object has a then function that throws if invoked$/, run: c => { const s = state(c); s.thenCalls = 0; s.thenable = { then: () => { s.thenCalls!++; throw new Error('must not be invoked'); } }; } },
  { pattern: /^a synchronous transaction replaces Alpha with Beta and returns that thenable object$/, run: c => { const s = state(c); assert(s.thenable); s.transactionResult = pkg(s).transaction(() => { pkg(s).set(doc, pkg(s).text(doc).replace('Alpha', 'Beta')); return s.thenable; }); } },
  { pattern: /^the transaction returns the same object by identity without invoking then$/, run: c => { const s = state(c); assert(s.thenable); assert.equal(s.transactionResult, s.thenable); assert.equal(s.thenCalls, 0); } },
  { pattern: /^the document text contains Beta$/, run: async c => { const s = state(c); assert(pkg(s).text(doc).includes('Beta')); s.saved = pkg(s).toBytes(); savedCustody(s); assert((await OpcPackage.open(s.saved)).text(doc).includes('Beta')); assert.equal(s.thenCalls, 0); } },
  { pattern: /^(?:an existing|a regular) destination file contains the base package archive bytes$/, run: async c => { await destination(state(c)); } },
  { pattern: /^the package editor has opened those bytes and deleted word\/document.xml$/, run: async c => { const s = state(c); await open(s); pkg(s).delete(doc); assert.equal(pkg(s).get(doc), undefined); } },
  { pattern: /^a symlink points to that file$/, run: async c => { const s = state(c); assert(s.destination); s.link = s.destination + '.link'; await symlink(s.destination, s.link); } },
  { pattern: /^(?:the package is saved to the existing destination|the unchanged package is saved through the symlink path)$/, run: async c => {
    const s = state(c); assert(s.destination); retainSource(s);
    try { s.output = await pkg(s).save(s.link ?? s.destination); } catch (error) { s.error = error; }
  } },
  {pattern:/^the symlink still points to the original regular destination without a successful save receipt$/,run:async c=>{
    const s=state(c);assert(s.error);assert.equal(s.output,undefined);assert(s.link&&s.destination);assert((await lstat(s.link)).isSymbolicLink());assert.equal(await readlink(s.link),s.destination);
  }},
  { pattern: /^the (?:regular )?destination file bytes equal the original archive bytes$/, run: async c => {
    const s = state(c); assert(s.destination && s.original); assert.deepEqual(Uint8Array.from(await Bun.file(s.destination).bytes()), s.original);
    assert.deepEqual((await OpcPackage.open(s.destination)).toBytes(), s.original);
    if (s.link) { assert((await lstat(s.link)).isSymbolicLink()); assert.equal(await readlink(s.link), s.destination); }
  } },
];
