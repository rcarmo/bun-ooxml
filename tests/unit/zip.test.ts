import {fixturePaths,CORPUS74_PACKAGE_IDS} from "../../scripts/fixture-inputs.ts";
import { describe, expect, test } from "bun:test";
import { buildZip, concatBytes } from "../fixtures/zip32.ts";
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, resolve } from "node:path";

import { OoxmlError } from "../../src/errors.ts";
import { crc32, readZip, writeZip } from "../../src/opc/zip.ts";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const EXPECTED_FIXTURE_ARCHIVE_COUNT = 74;

const FLAG_UTF8 = 0x0800;
const FLAG_DATA_DESCRIPTOR = 0x0008;
const METHOD_STORED = 0;
const METHOD_DEFLATED = 8;

describe("crc32", () => {
  test("matches the standard test vector", () => {
    expect(crc32(encoder.encode("123456789"))).toBe(0xcbf43926);
  });
});

describe("writeZip", () => {
  test("writes deterministic UTF-8 ZIP32 output that readZip can reopen", () => {
    const parts = new Map<string, Uint8Array>([
      ["word/", new Uint8Array()],
      ["word/document.xml", encoder.encode("<w:document/>")],
      ["[Content_Types].xml", encoder.encode("<Types/>")],
      ["docProps/core.xml", encoder.encode("<core/>")],
    ]);

    const first = writeZip(parts);
    const second = writeZip(parts);

    expect(first).toEqual(second);

    const reopened = readZip(first);
    expect(Array.from(reopened.keys())).toEqual([
      "[Content_Types].xml",
      "docProps/core.xml",
      "word/document.xml",
    ]);
    expect(decoder.decode(reopened.get("word/document.xml")!)).toBe("<w:document/>");
    expect(decoder.decode(reopened.get("[Content_Types].xml")!)).toBe("<Types/>");
  });

  test("refuses ASCII case-colliding output names", () => {
    const parts = new Map<string, Uint8Array>([
      ["word/document.xml", encoder.encode("one")],
      ["WORD/document.xml", encoder.encode("two")],
    ]);

    { const error = captureZipError(() => writeZip(parts), "zip-case-collision");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-case-collision");
    expect((error as OoxmlError).message).toContain("ASCII case-collides"); }
  });

  test("refuses non-empty directory entries", () => {
    const parts = new Map<string, Uint8Array>([["word/", encoder.encode("not empty")]]);

    { const error = captureZipError(() => writeZip(parts), "zip-directory-entry-invalid");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-directory-entry-invalid");
    expect((error as OoxmlError).message).toContain("must be empty"); }
  });
});

describe("readZip", () => {
  test("reads a valid archive with a directory entry, data descriptor, and declared comment", () => {
    const archive = buildZip(
      [
        { name: "word/", method: METHOD_STORED, blob: new Uint8Array() },
        { name: "[Content_Types].xml", method: METHOD_STORED, blob: encoder.encode("<Types/>") },
        {
          name: "word/document.xml",
          method: METHOD_DEFLATED,
          blob: encoder.encode("<w:document/>"),
          flags: FLAG_UTF8 | FLAG_DATA_DESCRIPTOR,
          dataDescriptor: true,
          descriptorSignature: true,
        },
      ],
      { archiveComment: encoder.encode("kept as declared ZIP comment") },
    );

    const parts = readZip(archive);
    expect(Array.from(parts.keys())).toEqual(["[Content_Types].xml", "word/document.xml"]);
    expect(decoder.decode(parts.get("[Content_Types].xml")!)).toBe("<Types/>");
    expect(decoder.decode(parts.get("word/document.xml")!)).toBe("<w:document/>");
  });

  test("refuses duplicate member names", () => {
    const archive = buildZip([
      { name: "word/document.xml", blob: encoder.encode("one") },
      { name: "word/document.xml", blob: encoder.encode("two") },
    ]);

    { const error = captureZipError(() => readZip(archive), "zip-duplicate-entry");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-duplicate-entry");
    expect((error as OoxmlError).message).toContain("duplicate ZIP member"); }
  });

  test("refuses ASCII case-colliding member names", () => {
    const archive = buildZip([
      { name: "word/document.xml", blob: encoder.encode("one") },
      { name: "WORD/document.xml", blob: encoder.encode("two") },
    ]);

    { const error = captureZipError(() => readZip(archive), "zip-case-collision");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-case-collision");
    expect((error as OoxmlError).message).toContain("ASCII case-collides"); }
  });

  test("refuses noncanonical part paths", () => {
    const archive = buildZip([{ name: "../word/document.xml", blob: encoder.encode("bad") }]);

    { const error = captureZipError(() => readZip(archive), "zip-name-invalid");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-name-invalid");
    expect((error as OoxmlError).message).toContain("noncanonical"); }
  });

  test("refuses encrypted members", () => {
    const archive = buildZip([
      { name: "word/document.xml", blob: encoder.encode("secret"), flags: FLAG_UTF8 | 0x0001 },
    ]);

    { const error = captureZipError(() => readZip(archive), "zip-encryption-unsupported");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-encryption-unsupported");
    expect((error as OoxmlError).message).toContain("encrypted"); }
  });

  test("refuses unsupported compression methods", () => {
    const archive = buildZip([{ name: "word/document.xml", blob: encoder.encode("x"), method: 12 }]);

    { const error = captureZipError(() => readZip(archive), "zip-method-unsupported");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-method-unsupported");
    expect((error as OoxmlError).message).toContain("compression method"); }
  });

  test("refuses multi-disk archives", () => {
    const archive = buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }], {
      diskNumber: 1,
    });

    { const error = captureZipError(() => readZip(archive), "zip-multi-disk-unsupported");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-multi-disk-unsupported");
    expect((error as OoxmlError).message).toContain("multi-disk"); }
  });

  test("refuses malformed ZIP64 sentinels without ZIP64 directory records", () => {
    const archive = buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }], {
      totalEntries: 0xffff,
      diskEntries: 0xffff,
    });

    { const error = captureZipError(() => readZip(archive), "zip-structure-invalid");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-structure-invalid");
    expect((error as OoxmlError).message).toContain("ZIP64"); }
  });

  test("refuses local and central metadata mismatches", () => {
    const archive = buildZip([
      {
        name: "word/document.xml",
        localName: "word/other.xml",
        blob: encoder.encode("x"),
      },
    ]);

    { const error = captureZipError(() => readZip(archive), "zip-local-metadata-mismatch");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-local-metadata-mismatch");
    expect((error as OoxmlError).message).toContain("local and central"); }
  });

  test("refuses CRC mismatches", () => {
    const archive = buildZip([
      {
        name: "word/document.xml",
        blob: encoder.encode("payload"),
        crc: 0xdeadbeef,
      },
    ]);

    { const error = captureZipError(() => readZip(archive), "zip-crc-mismatch");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-crc-mismatch");
    expect((error as OoxmlError).message).toContain("CRC"); }
  });

  test("refuses declared size mismatches", () => {
    const archive = buildZip([
      {
        name: "word/document.xml",
        blob: encoder.encode("payload"),
        method: METHOD_STORED,
        fileSize: 99,
        compressedSize: 99,
        localFileSize: 99,
        localCompressedSize: 99,
      },
    ]);

    { const error = captureZipError(() => readZip(archive), "zip-size-mismatch");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-size-mismatch");
    expect((error as OoxmlError).message).toContain("declared size"); }
  });

  test("refuses deflated members that inflate past their declared size bound", () => {
    const archive = buildZip([
      {
        name: "word/document.xml",
        blob: encoder.encode("A".repeat(4096)),
        method: METHOD_DEFLATED,
        fileSize: 32,
        localFileSize: 32,
      },
    ]);

    { const error = captureZipError(() => readZip(archive), "zip-size-mismatch");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-size-mismatch");
    expect((error as OoxmlError).message).toContain("declared size"); }
  });

  test("refuses undeclared trailing bytes after the archive", () => {
    const archive = concatBytes(buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }]), encoder.encode("\n"));

    { const error = captureZipError(() => readZip(archive), "zip-end-record-missing");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-end-record-missing");
    expect((error as OoxmlError).message).toContain("end-of-central-directory"); }
  });

  test("enforces configured archive, count, entry, total, and ratio bounds", () => {
    const repetitive = encoder.encode("A".repeat(4096));
    const archive = buildZip([
      { name: "a.bin", blob: repetitive },
      { name: "b.bin", blob: encoder.encode("bb") },
    ]);

    { const error = captureZipError(() => readZip(archive, { maxArchiveBytes: archive.length - 1 }), "zip-archive-too-large");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-archive-too-large");
    expect((error as OoxmlError).message).toContain("archive bytes"); }
    { const error = captureZipError(() => readZip(archive, { maxEntries: 1 }), "zip-too-many-entries");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-too-many-entries");
    expect((error as OoxmlError).message).toContain("entry limit"); }
    { const error = captureZipError(() => readZip(archive, { maxEntryBytes: 8 }), "zip-entry-too-large");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-entry-too-large");
    expect((error as OoxmlError).message).toContain("entry limit"); }
    { const error = captureZipError(() => readZip(archive, { maxTotalBytes: 8 }), "zip-total-too-large");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-total-too-large");
    expect((error as OoxmlError).message).toContain("total expanded"); }
    { const error = captureZipError(() => readZip(archive, { maxCompressionRatio: 2 }), "zip-compression-ratio-exceeded");
    expect(error).toBeInstanceOf(OoxmlError);
    expect((error as OoxmlError).code).toBe("zip-compression-ratio-exceeded");
    expect((error as OoxmlError).message).toContain("compression ratio"); }
  });

  test("reads every checked-in rcarmo fixture archive under default bounds", () => {
    const fixtures = fixtureArchives();
    expect(fixtures).toHaveLength(EXPECTED_FIXTURE_ARCHIVE_COUNT);

    for (const file of fixtures) {
      const bytes = new Uint8Array(readFileSync(file));
      try {
        readZip(bytes);
      } catch (error) {
        throw new Error(`fixture ${file} failed ZIP read: ${(error as Error).message}`);
      }
    }
  });
});

function captureZipError(action: () => unknown, code: string): unknown {
  try {
    action();
    throw new Error(`expected ${code}`);
  } catch (error) {
    if (error instanceof Error && error.message === `expected ${code}`) throw error;
    return error;
  }
}

function fixtureArchives(): string[] { return fixturePaths(CORPUS74_PACKAGE_IDS); }
