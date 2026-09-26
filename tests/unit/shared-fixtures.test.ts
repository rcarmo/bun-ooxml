import {fixturePath,fixturesRoot} from "../../scripts/fixture-inputs.ts";
import { afterEach, describe, expect, test } from "bun:test";
import { existsSync } from "node:fs";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { Document, Presentation, Workbook } from "../../src/index.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import {
  assertPreservedOutput,
  type FixtureRecord,
  verifyFixture,
} from "../../scripts/shared-fixtures.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const PACK_ROOT = resolveSharedPackRoot();
const tempRoots: string[] = [];
const encoder = new TextEncoder();
const decoder = new TextDecoder("utf-8", { fatal: true });

afterEach(async () => {
  await Promise.all(tempRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("shared fixture validation", () => {
  test("verifies the four pinned shared fixtures and their readback preconditions", async () => {
    const manifest = await loadManifest();

    expect(manifest.fixtures.map((fixture) => fixture.id).sort()).toEqual([
      "cross-sheet-cache.xlsx",
      "default-style.xlsx",
      "present-placeholder.docx",
      "title-and-subtitle.pptx",
    ]);

    for (const fixture of manifest.fixtures) {
      await expect(verifyFixture(await readFixtureBytes(fixture), fixture)).resolves.toBeUndefined();
    }
  });

  test("verifyFixture rejects a wrong whole-archive hash before edit semantics", async () => {
    const fixture = await fixtureById("present-placeholder.docx");
    const mutated = rewritePart(
      await readFixtureBytes(fixture),
      "word/document.xml",
      (xml) => xml.replace("&lt;Present&gt;", "&lt;Changed&gt;"),
    );

    await expect(verifyFixture(mutated, fixture)).rejects.toThrow(/archive sha256/i);
  });

  test("assertPreservedOutput rejects opaque member drift outside the allowed parts", async () => {
    const fixture = await fixtureById("title-and-subtitle.pptx");
    const mutated = rewritePart(
      await readFixtureBytes(fixture),
      "docProps/app.xml",
      (xml) => xml.replace("<Application>", "<Application data-test=\"changed\">") ,
    );

    await expect(assertPreservedOutput(mutated, fixture)).rejects.toThrow(/unexpected payload drift/i);
  });

  test("assertPreservedOutput rejects a detached sentinel even if the root rel bytes are re-pinned", async () => {
    const fixture = await fixtureById("default-style.xlsx");
    const mutated = rewritePart(
      await readFixtureBytes(fixture),
      "_rels/.rels",
      (xml) => xml.replace('Target="customXml/preservation-sentinel.xml"', 'Target="docProps/app.xml"'),
    );
    const retargeted = withPinnedMember(fixture, "_rels/.rels", partBytes(mutated, "_rels/.rels"));

    await expect(assertPreservedOutput(mutated, retargeted)).rejects.toThrow(/sentinel root relationship/i);
  });

  test("assertPreservedOutput rejects invalid XLSX style indices in allowed changed worksheet parts", async () => {
    const fixture = await fixtureById("default-style.xlsx");
    const mutated = rewritePart(
      await readFixtureBytes(fixture),
      "xl/worksheets/sheet1.xml",
      (xml) => xml.replace('<s:c r="A1" t="inlineStr">', '<s:c r="A1" s="99" t="inlineStr">'),
    );

    await expect(assertPreservedOutput(mutated, fixture)).rejects.toThrow(/style index/i);
  });

  test("implicit style zero still requires a cellXfs definition", async () => {
    const fixture = await fixtureById("default-style.xlsx");
    const mutated = rewritePart(await readFixtureBytes(fixture), "xl/styles.xml",
      xml => xml.replace(/<cellXfs\b[^>]*>[\s\S]*?<\/cellXfs>/, '<cellXfs count="0"/>'));
    await expect(assertPreservedOutput(mutated, fixture)).rejects.toThrow(/style index/i);
  });

  test("assertPreservedOutput rejects unexpected added members", async () => {
    const fixture = await fixtureById("cross-sheet-cache.xlsx");
    const bytes = await readFixtureBytes(fixture);
    const parts = readZip(bytes);
    parts.set("customXml/unexpected.xml", encoder.encode("<unexpected/>"));

    await expect(assertPreservedOutput(writeZip(parts), fixture)).rejects.toThrow(/unexpected added parts/i);
  });

  test("assertPreservedOutput rejects unexpected deleted members", async () => {
    const fixture = await fixtureById("present-placeholder.docx");
    const bytes = await readFixtureBytes(fixture);
    const parts = readZip(bytes);
    parts.delete("customXml/preservation-sentinel.xml");

    await expect(assertPreservedOutput(writeZip(parts), fixture)).rejects.toThrow(/unexpected deleted parts/i);
  });
});

describe("shared fixture preservation on supported native edits", () => {
  test("DOCX edit preserves shared fixture invariants after save and reopen (unit-only)", async () => {
    const fixture = await fixtureById("present-placeholder.docx");
    const root = await tempRoot();
    const output = join(root, fixture.id);
    const document = await Document.open(await readFixtureBytes(fixture));
    const match = document.find("<Present>")[0];
    if (!match) throw new Error("expected DOCX match");

    await match.replace("<Present now>");
    await document.save(output);

    const savedBytes = Uint8Array.from(await Bun.file(output).bytes());
    const reopened = await Document.open(output);

    expect(reopened.paragraphs.map((paragraph) => paragraph.text)).toEqual(["<Present now>"]);
    await expect(assertPreservedOutput(savedBytes, fixture)).resolves.toBeUndefined();
  });

  test("PPTX edit preserves shared fixture invariants after save and reopen (unit-only)", async () => {
    const fixture = await fixtureById("title-and-subtitle.pptx");
    const root = await tempRoot();
    const output = join(root, fixture.id);
    const presentation = await Presentation.open(await readFixtureBytes(fixture));
    const paragraph = presentation.slides[0]?.inspectText("shared-fixture.unit")[0];
    if (!paragraph) throw new Error("expected PPTX title paragraph");

    presentation.slides[0]!.replaceTextAt(paragraph.anchor, "Original title", "Updated title");
    await presentation.save(output);

    const savedBytes = Uint8Array.from(await Bun.file(output).bytes());
    const reopened = await Presentation.open(output);

    expect(reopened.slides[0]?.inspectText("shared-fixture.unit").map((entry) => entry.text)).toEqual([
      "Updated title",
      "Original subtitle",
    ]);
    await expect(assertPreservedOutput(savedBytes, fixture)).resolves.toBeUndefined();
  });

  test("XLSX cache edit preserves shared fixture invariants after save and reopen (unit-only)", async () => {
    const fixture = await fixtureById("cross-sheet-cache.xlsx");
    const root = await tempRoot();
    const output = join(root, fixture.id);
    const workbook = await Workbook.open(await readFixtureBytes(fixture));

    workbook.worksheet("Input").setCellValue("A1", 10);
    await workbook.save(output);

    const savedBytes = Uint8Array.from(await Bun.file(output).bytes());
    const reopened = await Workbook.open(output);

    expect(reopened.worksheet("Input").getCell("A1")).toMatchObject({
      kind: "number",
      value: 10,
    });
    expect(reopened.worksheet("Calc").getCell("A1")).toMatchObject({
      kind: "formula",
      formula: "Input!A1*2",
      cached: null,
    });
    await expect(assertPreservedOutput(savedBytes, fixture)).resolves.toBeUndefined();
  });
});

type FixtureManifest = {
  schemaVersion: number;
  contractRevision: string;
  fixtures: FixtureRecord[];
};

async function loadManifest(): Promise<FixtureManifest> {
  return Bun.file(join(PACK_ROOT, "fixture-manifest.json")).json() as Promise<FixtureManifest>;
}

async function fixtureById(id: string): Promise<FixtureRecord> {
  const fixture = (await loadManifest()).fixtures.find((entry) => entry.id === id);
  if (!fixture) throw new Error(`Missing fixture ${id}`);
  return fixture;
}

async function readFixtureBytes(fixture: FixtureRecord): Promise<Uint8Array> {
  return Uint8Array.from(await Bun.file(fixturePath(fixture.assetId)).bytes());
}

function resolveSharedPackRoot(): string {
  const preferred = join(fixturesRoot(), "shared/v2/pack");
  if (!existsSync(join(preferred, "fixture-manifest.json"))) throw new Error("Missing pinned shared fixture pack");
  return preferred;
}

async function tempRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-shared-fixtures-"));
  tempRoots.push(root);
  return root;
}

function rewritePart(bytes: Uint8Array, name: string, rewrite: (text: string) => string): Uint8Array {
  const parts = readZip(bytes);
  const original = parts.get(name);
  if (!original) throw new Error(`Missing package part ${name}`);
  const next = rewrite(decoder.decode(original));
  if (next === decoder.decode(original)) throw new Error(`Expected ${name} to change`);
  parts.set(name, encoder.encode(next));
  return writeZip(parts);
}

function partBytes(bytes: Uint8Array, name: string): Uint8Array {
  const part = readZip(bytes).get(name);
  if (!part) throw new Error(`Missing package part ${name}`);
  return part;
}

function withPinnedMember(fixture: FixtureRecord, name: string, bytes: Uint8Array): FixtureRecord {
  const clone = structuredClone(fixture);
  const sha256 = new Bun.CryptoHasher("sha256").update(bytes).digest("hex");
  clone.memberSha256[name] = sha256;
  if (clone.mustPreservePayloads[name] !== undefined) {
    clone.mustPreservePayloads[name] = sha256;
  }
  return clone;
}
