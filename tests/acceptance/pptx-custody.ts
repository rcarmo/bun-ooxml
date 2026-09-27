import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { fixturePath, F } from '../../scripts/fixture-inputs.ts';
import { Presentation } from '../../src/pptx/index.ts';
export type CustodyState = { sourcePath: string; original: Uint8Array; fromPath?: Presentation; fromBytes?: Presentation; saved?: Uint8Array; reopened?: Uint8Array; destination?: string };
const state = (c: Record<string, unknown>) => c.state as CustodyState;
export const scenarioIds = ['@id-pptx-bun-open-save-noop'];
export const bindings: StepBinding[] = [
  { pattern: new RegExp('^fixture ' + F.officeSlides.titleSlide + '$'), run: async c => {
    const s = state(c); s.sourcePath = fixturePath(F.officeSlides.titleSlide); s.original = Uint8Array.from(await Bun.file(s.sourcePath).bytes());
  } },
  { pattern: /^the presentation reader opens the fixture path and separately opens its archive bytes$/, run: async c => {
    const s = state(c); assert(s.sourcePath && s.original); s.fromPath = await Presentation.open(s.sourcePath); s.fromBytes = await Presentation.open(s.original);
  } },
  { pattern: /^the path-opened presentation's first inspected paragraph on its first slide is (.+)$/, run: (c, text) => {
    const s = state(c); assert(s.fromPath); assert.equal(s.fromPath.slides[0]?.inspectText('acceptance.open.path')[0]?.text, text);
  } },
  { pattern: /^serializing the byte-opened presentation returns the exact original archive bytes$/, run: c => {
    const s = state(c); assert(s.fromBytes && s.original); assert.deepEqual(s.fromBytes.package.toBytes(), s.original);
  } },
  { pattern: /^the path-opened presentation is saved without edits to a new PPTX path$/, run: async c => {
    const s = state(c); assert(s.fromPath); const root = await mkdtemp(join(tmpdir(), 'bun-pptx-custody-'));
    try {
      s.destination = join(root, 'copy.pptx'); await s.fromPath.save(s.destination);
      s.saved = Uint8Array.from(await Bun.file(s.destination).bytes()); s.reopened = (await Presentation.open(s.destination)).package.toBytes();
    } finally { await rm(root, { recursive: true, force: true }); }
  } },
  { pattern: /^that destination file contains the exact original archive bytes$/, run: async c => {
    const s = state(c); assert(s.saved && s.reopened && s.original);
    assert.deepEqual(s.saved, s.original); assert.deepEqual(s.reopened, s.original);
    assert.deepEqual(Uint8Array.from(await Bun.file(s.sourcePath).bytes()), s.original);
  } },
];
