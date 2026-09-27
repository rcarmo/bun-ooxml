import {fixturePath,fixturePaths,F,CORPUS74_PACKAGE_IDS} from "../../scripts/fixture-inputs.ts";
import { afterEach, describe, expect, test } from "bun:test";
import { readdirSync, statSync } from "node:fs";
import { mkdtemp, rm, symlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

import { OoxmlError } from "../../src/errors.ts";
import { OpcPackage } from "../../src/opc/package.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";

const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const MAIN_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const DOCUMENT_PART = "word/document.xml";
const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const EXPECTED_FIXTURE_ARCHIVE_COUNT = 73;
const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });
const UTF16LE_DECODER = new TextDecoder("utf-16le", { fatal: true });
const UTF8_ENCODER = new TextEncoder();
const roots: string[] = [];

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("OpcPackage", () => {
  test("conservatively refuses percent-encoded internal names and relationship targets", async () => {
    { const error = await captureOpcErrorAsync(OpcPackage.open(createPackageBytes({ documentName: "word/%66oo.xml" })), "opc-part-name-invalid");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("opc-part-name-invalid");
    expect((error as OoxmlError).message).toContain("Noncanonical part name"); }

    { const error = await captureOpcErrorAsync(OpcPackage.open(createPackageBytes({ relTarget: "word/%66oo.xml" })), "opc-target-invalid");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("opc-target-invalid");
    expect((error as OoxmlError).message).toContain("Percent-encoded"); }
  });

  test("validates content-type metadata", async () => {
    { const error = await captureOpcErrorAsync(OpcPackage.open(createPackageBytes({
        contentTypesXml:
          `<?xml version="1.0" encoding="UTF-8"?>`
          + `<Types xmlns="${CONTENT_TYPES_NS}">`
          + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
          + `<Default Extension="rels" ContentType="application/xml"/>`
          + `<Override PartName="/${DOCUMENT_PART}" ContentType="${MAIN_CONTENT_TYPE}"/>`
          + `</Types>`,
      })), "opc-content-types-invalid");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("opc-content-types-invalid");
    expect((error as OoxmlError).message).toContain("Duplicate/missing default extension"); }
  });

  test("rejects missing and duplicate OPC relationship references", async () => {
    { const error = await captureOpcErrorAsync(OpcPackage.open(createPackageBytes({ relTarget: "word/missing.xml" })), "opc-relationship-target-missing");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("opc-relationship-target-missing");
    expect((error as OoxmlError).message).toContain("Missing target"); }

    { const error = await captureOpcErrorAsync(OpcPackage.open(createPackageBytes({
        rootRelationshipsXml:
          `<?xml version="1.0" encoding="UTF-8"?>`
          + `<Relationships xmlns="${RELATIONSHIPS_NS}">`
          + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="${DOCUMENT_PART}"/>`
          + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="${DOCUMENT_PART}"/>`
          + `</Relationships>`,
      })), "opc-relationship-duplicate");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("opc-relationship-duplicate");
    expect((error as OoxmlError).message).toContain("Duplicate relationship id"); }
  });

  test("returns detached copies and keeps custody of opened source bytes", async () => {
    const source = createPackageBytes();
    const original = source.slice();
    const pkg = await OpcPackage.open(source);

    source.fill(0);
    const first = pkg.get(DOCUMENT_PART)!;
    first.fill(0);

    expect(UTF8_DECODER.decode(pkg.get(DOCUMENT_PART)!)).toContain("Alpha");
    expect(pkg.toBytes()).toEqual(original);
  });

  test("preserves UTF-16 XML encoding when string edits are written back", async () => {
    const utf16Document = encodeUtf16Le(
      `<?xml version="1.0" encoding="UTF-16"?><document>Alpha</document>`,
    );
    const pkg = await OpcPackage.open(createPackageBytes({ documentBytes: utf16Document }));

    pkg.set(DOCUMENT_PART, pkg.text(DOCUMENT_PART).replace("Alpha", "Beta"));

    const saved = readZip(pkg.toBytes()).get(DOCUMENT_PART)!;
    expect([...saved.slice(0, 2)]).toEqual([0xff, 0xfe]);
    expect(UTF16LE_DECODER.decode(saved)).toContain("Beta");
    expect(UTF16LE_DECODER.decode(saved)).toContain('encoding="UTF-16"');
  });

  test("rejects async transaction callbacks before they run and does not special-case thenables", async () => {
    const pkg = await OpcPackage.open(createPackageBytes());
    let ran = false;

    { const error = captureOpcErrorSync(() => pkg.transaction(async () => {
        ran = true;
        pkg.set(DOCUMENT_PART, `<?xml version="1.0" encoding="UTF-8"?><document>Beta</document>`);
      }), "opc-async-transaction");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("opc-async-transaction");
    expect((error as OoxmlError).message).toContain("synchronous edits"); }

    expect(ran).toBe(false);
    expect(pkg.text(DOCUMENT_PART)).toContain("Alpha");

    const thenable = { then: () => { throw new Error("must not be awaited"); } };
    const result = pkg.transaction(() => {
      pkg.set(DOCUMENT_PART, `<?xml version="1.0" encoding="UTF-8"?><document>Beta</document>`);
      return thenable;
    });

    expect(result).toBe(thenable);
    expect(pkg.text(DOCUMENT_PART)).toContain("Beta");
  });

  test("save refuses invalid packages without touching an existing target", async () => {
    const bytes = createPackageBytes();
    const root = await tempRoot();
    const path = join(root, "existing.docx");
    await Bun.write(path, bytes);

    const pkg = await OpcPackage.open(bytes);
    pkg.delete(DOCUMENT_PART);

    { const error = await captureOpcErrorAsync(pkg.save(path), "opc-relationship-target-missing");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("opc-relationship-target-missing");
    expect((error as OoxmlError).message).toContain("Missing target"); }
    expect([...await Bun.file(path).bytes()]).toEqual([...bytes]);
  });

  test("refuses symlink save destinations", async () => {
    const bytes = createPackageBytes();
    const root = await tempRoot();
    const target = join(root, "target.docx");
    const link = join(root, "target-link.docx");
    await Bun.write(target, bytes);
    await symlink(target, link);

    const pkg = await OpcPackage.open(bytes);

    { const error = await captureOpcErrorAsync(pkg.save(link), "opc-symlink-destination");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("opc-symlink-destination");
    expect((error as OoxmlError).message).toContain("Symlink"); }
    expect([...await Bun.file(target).bytes()]).toEqual([...bytes]);
  });

  test("preserves exact bytes for every no-op fixture archive in the rcarmo corpora", async () => {
    const fixtures = fixtureArchives();
    expect(fixtures).toHaveLength(EXPECTED_FIXTURE_ARCHIVE_COUNT);

    for (const path of fixtures) {
      const original = Uint8Array.from(await Bun.file(path).bytes());
      const pkg = await OpcPackage.open(original);
      expect(pkg.toBytes()).toEqual(original);
    }
  });
});

function createPackageBytes(options: {
  documentName?: string;
  documentBytes?: Uint8Array;
  relTarget?: string;
  rootRelationshipsXml?: string;
  contentTypesXml?: string;
} = {}): Uint8Array {
  const documentName = options.documentName ?? DOCUMENT_PART;
  return writeZip(new Map<string, Uint8Array>([
    ["[Content_Types].xml", UTF8_ENCODER.encode(options.contentTypesXml ?? contentTypesXml(documentName))],
    ["_rels/.rels", UTF8_ENCODER.encode(options.rootRelationshipsXml ?? relationshipsXml(options.relTarget ?? documentName))],
    [documentName, options.documentBytes ?? UTF8_ENCODER.encode(`<?xml version="1.0" encoding="UTF-8"?><document>Alpha</document>`)],
  ]));
}

function contentTypesXml(documentName: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>`
    + `<Types xmlns="${CONTENT_TYPES_NS}">`
    + `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
    + `<Default Extension="xml" ContentType="application/xml"/>`
    + `<Override PartName="/${documentName}" ContentType="${MAIN_CONTENT_TYPE}"/>`
    + `</Types>`;
}

function relationshipsXml(target: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>`
    + `<Relationships xmlns="${RELATIONSHIPS_NS}">`
    + `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL}" Target="${target}"/>`
    + `</Relationships>`;
}

function encodeUtf16Le(text: string): Uint8Array {
  const bytes = new Uint8Array(2 + text.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index);
    const offset = 2 + index * 2;
    bytes[offset] = code & 0xff;
    bytes[offset + 1] = code >>> 8;
  }
  return bytes;
}

function fixtureArchives(): string[] { return fixturePaths(CORPUS74_PACKAGE_IDS); }

function captureOpcErrorSync(action: () => unknown, code: string): unknown {
  try { action(); throw new Error(`expected ${code}`); }
  catch (error) { if (error instanceof Error && error.message === `expected ${code}`) throw error; return error; }
}
async function captureOpcErrorAsync(action: Promise<unknown>, code: string): Promise<unknown> {
  try { await action; throw new Error(`expected ${code}`); }
  catch (error) { if (error instanceof Error && error.message === `expected ${code}`) throw error; return error; }
}

async function tempRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-opc-"));
  roots.push(root);
  return root;
}
