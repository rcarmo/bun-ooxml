import { describe, expect, test } from "bun:test";
import { deflateRawSync } from "node:zlib";
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, resolve } from "node:path";

import { OoxmlError } from "../../src/errors.ts";
import { crc32, readZip, writeZip } from "../../src/opc/zip.ts";

const encoder = new TextEncoder();
const decoder = new TextDecoder();
const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const FIXTURE_ROOTS = [
  join(PROJECT_ROOT, "fixtures/go-ooxml/testdata"),
  join(PROJECT_ROOT, "fixtures/python-office-mcp-server/tests/_templates"),
] as const;
const EXPECTED_FIXTURE_ARCHIVE_COUNT = 74;

const LOCAL_SIGNATURE = 0x04034b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;

const FLAG_UTF8 = 0x0800;
const FLAG_DATA_DESCRIPTOR = 0x0008;

const METHOD_STORED = 0;
const METHOD_DEFLATED = 8;

type ZipMemberSpec = {
  name: string;
  blob?: Uint8Array;
  method?: number;
  flags?: number;
  data?: Uint8Array;
  crc?: number;
  fileSize?: number;
  compressedSize?: number;
  localName?: string;
  localFlags?: number;
  localMethod?: number;
  localCrc?: number;
  localFileSize?: number;
  localCompressedSize?: number;
  localExtra?: Uint8Array;
  centralExtra?: Uint8Array;
  comment?: Uint8Array;
  headerOffset?: number;
  diskStart?: number;
  dataDescriptor?: boolean;
  descriptorSignature?: boolean;
  descriptorCrc?: number;
  descriptorFileSize?: number;
  descriptorCompressedSize?: number;
};

type ZipBuildOptions = {
  diskNumber?: number;
  centralDisk?: number;
  diskEntries?: number;
  totalEntries?: number;
  centralSizeDelta?: number;
  archiveComment?: Uint8Array;
};

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

    expectZipError(
      () => writeZip(parts),
      "zip-case-collision",
      "ASCII case-collides",
    );
  });

  test("refuses non-empty directory entries", () => {
    const parts = new Map<string, Uint8Array>([["word/", encoder.encode("not empty")]]);

    expectZipError(
      () => writeZip(parts),
      "zip-directory-entry-invalid",
      "must be empty",
    );
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

    expectZipError(() => readZip(archive), "zip-duplicate-entry", "duplicate ZIP member");
  });

  test("refuses ASCII case-colliding member names", () => {
    const archive = buildZip([
      { name: "word/document.xml", blob: encoder.encode("one") },
      { name: "WORD/document.xml", blob: encoder.encode("two") },
    ]);

    expectZipError(() => readZip(archive), "zip-case-collision", "ASCII case-collides");
  });

  test("refuses noncanonical part paths", () => {
    const archive = buildZip([{ name: "../word/document.xml", blob: encoder.encode("bad") }]);

    expectZipError(() => readZip(archive), "zip-name-invalid", "noncanonical");
  });

  test("refuses encrypted members", () => {
    const archive = buildZip([
      { name: "word/document.xml", blob: encoder.encode("secret"), flags: FLAG_UTF8 | 0x0001 },
    ]);

    expectZipError(() => readZip(archive), "zip-encryption-unsupported", "encrypted");
  });

  test("refuses unsupported compression methods", () => {
    const archive = buildZip([{ name: "word/document.xml", blob: encoder.encode("x"), method: 12 }]);

    expectZipError(() => readZip(archive), "zip-method-unsupported", "compression method");
  });

  test("refuses multi-disk archives", () => {
    const archive = buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }], {
      diskNumber: 1,
    });

    expectZipError(() => readZip(archive), "zip-multi-disk-unsupported", "multi-disk");
  });

  test("refuses malformed ZIP64 sentinels without ZIP64 directory records", () => {
    const archive = buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }], {
      totalEntries: 0xffff,
      diskEntries: 0xffff,
    });

    expectZipError(() => readZip(archive), "zip-structure-invalid", "ZIP64");
  });

  test("refuses local and central metadata mismatches", () => {
    const archive = buildZip([
      {
        name: "word/document.xml",
        localName: "word/other.xml",
        blob: encoder.encode("x"),
      },
    ]);

    expectZipError(() => readZip(archive), "zip-local-metadata-mismatch", "local and central");
  });

  test("refuses CRC mismatches", () => {
    const archive = buildZip([
      {
        name: "word/document.xml",
        blob: encoder.encode("payload"),
        crc: 0xdeadbeef,
      },
    ]);

    expectZipError(() => readZip(archive), "zip-crc-mismatch", "CRC");
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

    expectZipError(() => readZip(archive), "zip-size-mismatch", "declared size");
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

    expectZipError(() => readZip(archive), "zip-size-mismatch", "declared size");
  });

  test("refuses undeclared trailing bytes after the archive", () => {
    const archive = concatBytes(buildZip([{ name: "word/document.xml", blob: encoder.encode("x") }]), encoder.encode("\n"));

    expectZipError(() => readZip(archive), "zip-end-record-missing", "end-of-central-directory");
  });

  test("enforces configured archive, count, entry, total, and ratio bounds", () => {
    const repetitive = encoder.encode("A".repeat(4096));
    const archive = buildZip([
      { name: "a.bin", blob: repetitive },
      { name: "b.bin", blob: encoder.encode("bb") },
    ]);

    expectZipError(() => readZip(archive, { maxArchiveBytes: archive.length - 1 }), "zip-archive-too-large", "archive bytes");
    expectZipError(() => readZip(archive, { maxEntries: 1 }), "zip-too-many-entries", "entry limit");
    expectZipError(() => readZip(archive, { maxEntryBytes: 8 }), "zip-entry-too-large", "entry limit");
    expectZipError(() => readZip(archive, { maxTotalBytes: 8 }), "zip-total-too-large", "total expanded");
    expectZipError(() => readZip(archive, { maxCompressionRatio: 2 }), "zip-compression-ratio-exceeded", "compression ratio");
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

function expectZipError(action: () => unknown, code: string, messageFragment: string): void {
  try {
    action();
    throw new Error(`expected ${code}`);
  } catch (error) {
    if (error instanceof Error && error.message === `expected ${code}`) {
      throw error;
    }
    expect(error).toBeInstanceOf(OoxmlError);
    const refusal = error as OoxmlError;
    expect(refusal.code).toBe(code);
    expect(refusal.message).toContain(messageFragment);
  }
}

function fixtureArchives(): string[] {
  return FIXTURE_ROOTS.flatMap(collectFixtureArchives).sort();
}

function collectFixtureArchives(root: string): string[] {
  if (!statSync(root).isDirectory()) throw new Error(`Missing fixture root: ${root}`);

  const result: string[] = [];
  const visit = (dir: string): void => {
    for (const entry of readdirSync(dir, { withFileTypes: true })) {
      const next = join(dir, entry.name);
      if (entry.isDirectory()) {
        visit(next);
        continue;
      }
      if (/\.(zip|docx|pptx|xlsx)$/i.test(entry.name)) result.push(next);
    }
  };

  visit(root);
  return result;
}

function buildZip(members: ZipMemberSpec[], options: ZipBuildOptions = {}): Uint8Array {
  const bodyChunks: Uint8Array[] = [];
  const centralChunks: Uint8Array[] = [];
  let bodyLength = 0;

  for (const spec of members) {
    const blob = spec.blob ?? encoder.encode(spec.name);
    const method = spec.method ?? METHOD_DEFLATED;
    const flags = spec.flags ?? FLAG_UTF8;
    const compressed = spec.data ?? compress(method, blob);
    const crc = spec.crc ?? crc32Test(blob);
    const fileSize = spec.fileSize ?? blob.length;
    const compressedSize = spec.compressedSize ?? compressed.length;
    const localFlags = spec.localFlags ?? flags;
    const localMethod = spec.localMethod ?? method;
    const localCrc = spec.localCrc ?? (flags & FLAG_DATA_DESCRIPTOR ? 0 : crc);
    const localFileSize = spec.localFileSize ?? (flags & FLAG_DATA_DESCRIPTOR ? 0 : fileSize);
    const localCompressedSize =
      spec.localCompressedSize ?? (flags & FLAG_DATA_DESCRIPTOR ? 0 : compressedSize);
    const rawCentralName = encoder.encode(spec.name);
    const rawLocalName = encoder.encode(spec.localName ?? spec.name);
    const localExtra = spec.localExtra ?? new Uint8Array();
    const centralExtra = spec.centralExtra ?? new Uint8Array();
    const comment = spec.comment ?? new Uint8Array();
    const headerOffset = spec.headerOffset ?? bodyLength;
    const descriptor = spec.dataDescriptor
      ? buildDataDescriptor({
          crc: spec.descriptorCrc ?? crc,
          compressedSize: spec.descriptorCompressedSize ?? compressedSize,
          fileSize: spec.descriptorFileSize ?? fileSize,
          withSignature: spec.descriptorSignature ?? false,
        })
      : new Uint8Array();

    bodyChunks.push(
      buildLocalHeader({
        flags: localFlags,
        method: localMethod,
        crc: localCrc,
        compressedSize: localCompressedSize,
        fileSize: localFileSize,
        nameLength: rawLocalName.length,
        extraLength: localExtra.length,
      }),
      rawLocalName,
      localExtra,
      compressed,
      descriptor,
    );
    bodyLength +=
      30 + rawLocalName.length + localExtra.length + compressed.length + descriptor.length;

    centralChunks.push(
      buildCentralHeader({
        flags,
        method,
        crc,
        compressedSize,
        fileSize,
        nameLength: rawCentralName.length,
        extraLength: centralExtra.length,
        commentLength: comment.length,
        diskStart: spec.diskStart ?? 0,
        headerOffset,
      }),
      rawCentralName,
      centralExtra,
      comment,
    );
  }

  const centralOffset = bodyLength;
  const centralDirectory = concatBytes(...centralChunks);
  const archiveComment = options.archiveComment ?? new Uint8Array();
  const endRecord = buildEndRecord({
    diskNumber: options.diskNumber ?? 0,
    centralDisk: options.centralDisk ?? 0,
    diskEntries: options.diskEntries ?? members.length,
    totalEntries: options.totalEntries ?? members.length,
    centralSize: centralDirectory.length + (options.centralSizeDelta ?? 0),
    centralOffset,
    commentLength: archiveComment.length,
  });

  return concatBytes(...bodyChunks, centralDirectory, endRecord, archiveComment);
}

function buildLocalHeader(fields: {
  flags: number;
  method: number;
  crc: number;
  compressedSize: number;
  fileSize: number;
  nameLength: number;
  extraLength: number;
}): Uint8Array {
  const bytes = new Uint8Array(30);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, LOCAL_SIGNATURE, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, fields.flags, true);
  view.setUint16(8, fields.method, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, 33, true);
  view.setUint32(14, fields.crc >>> 0, true);
  view.setUint32(18, fields.compressedSize >>> 0, true);
  view.setUint32(22, fields.fileSize >>> 0, true);
  view.setUint16(26, fields.nameLength, true);
  view.setUint16(28, fields.extraLength, true);
  return bytes;
}

function buildCentralHeader(fields: {
  flags: number;
  method: number;
  crc: number;
  compressedSize: number;
  fileSize: number;
  nameLength: number;
  extraLength: number;
  commentLength: number;
  diskStart: number;
  headerOffset: number;
}): Uint8Array {
  const bytes = new Uint8Array(46);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, CENTRAL_SIGNATURE, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, fields.flags, true);
  view.setUint16(10, fields.method, true);
  view.setUint16(12, 0, true);
  view.setUint16(14, 33, true);
  view.setUint32(16, fields.crc >>> 0, true);
  view.setUint32(20, fields.compressedSize >>> 0, true);
  view.setUint32(24, fields.fileSize >>> 0, true);
  view.setUint16(28, fields.nameLength, true);
  view.setUint16(30, fields.extraLength, true);
  view.setUint16(32, fields.commentLength, true);
  view.setUint16(34, fields.diskStart, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, fields.headerOffset >>> 0, true);
  return bytes;
}

function buildEndRecord(fields: {
  diskNumber: number;
  centralDisk: number;
  diskEntries: number;
  totalEntries: number;
  centralSize: number;
  centralOffset: number;
  commentLength: number;
}): Uint8Array {
  const bytes = new Uint8Array(22);
  const view = new DataView(bytes.buffer);
  view.setUint32(0, END_SIGNATURE, true);
  view.setUint16(4, fields.diskNumber, true);
  view.setUint16(6, fields.centralDisk, true);
  view.setUint16(8, fields.diskEntries, true);
  view.setUint16(10, fields.totalEntries, true);
  view.setUint32(12, fields.centralSize >>> 0, true);
  view.setUint32(16, fields.centralOffset >>> 0, true);
  view.setUint16(20, fields.commentLength, true);
  return bytes;
}

function buildDataDescriptor(fields: {
  crc: number;
  compressedSize: number;
  fileSize: number;
  withSignature: boolean;
}): Uint8Array {
  const size = fields.withSignature ? 16 : 12;
  const bytes = new Uint8Array(size);
  const view = new DataView(bytes.buffer);
  let offset = 0;
  if (fields.withSignature) {
    view.setUint32(0, DATA_DESCRIPTOR_SIGNATURE, true);
    offset = 4;
  }
  view.setUint32(offset, fields.crc >>> 0, true);
  view.setUint32(offset + 4, fields.compressedSize >>> 0, true);
  view.setUint32(offset + 8, fields.fileSize >>> 0, true);
  return bytes;
}

function compress(method: number, blob: Uint8Array): Uint8Array {
  if (method === METHOD_STORED) {
    return blob;
  }
  if (method === METHOD_DEFLATED) {
    return new Uint8Array(deflateRawSync(Buffer.from(blob)));
  }
  return blob;
}

function concatBytes(...chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const merged = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }
  return merged;
}

function crc32Test(bytes: Uint8Array): number {
  let value = 0xffffffff;
  for (const byte of bytes) {
    value = CRC32_TABLE[(value ^ byte) & 0xff]! ^ (value >>> 8);
  }
  return (value ^ 0xffffffff) >>> 0;
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();
