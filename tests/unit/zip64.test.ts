import { describe, expect, test } from "bun:test";
import { deflateRawSync } from "node:zlib";

import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip, crc32 } from "../../src/opc/zip.ts";

const encoder = new TextEncoder();
const decoder = new TextDecoder();

const LOCAL_SIGNATURE = 0x04034b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
const ZIP64_END_SIGNATURE = 0x06064b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;
const ZIP64_EXTRA_TAG = 0x0001;

const FLAG_UTF8 = 0x0800;
const FLAG_DATA_DESCRIPTOR = 0x0008;

const METHOD_STORED = 0;
const METHOD_DEFLATED = 8;

const MAX_UINT16 = 0xffff;
const MAX_UINT32 = 0xffffffff;
const MAX_SAFE_BIGINT = BigInt(Number.MAX_SAFE_INTEGER);

type DescriptorKind = "none" | "32" | "32sig" | "64" | "64sig";

type Zip64MemberSpec = {
  name: string;
  blob?: Uint8Array;
  method?: number;
  dataDescriptor?: DescriptorKind;
  includeLocalZip64?: boolean;
  includeCentralZip64?: boolean;
  duplicateLocalZip64Extra?: boolean;
  duplicateCentralZip64Extra?: boolean;
};

type Zip64BuildOptions = {
  extensibleData?: Uint8Array;
};

type BuiltZip64Archive = {
  bytes: Uint8Array;
  centralOffset: number;
  centralSize: number;
  zip64EocdOffset: number;
  locatorOffset: number;
  eocdOffset: number;
  entries: Array<{
    name: string;
    localHeaderOffset: number;
    localExtraOffset: number;
    localExtraLength: number;
    dataOffset: number;
    dataLength: number;
    descriptorOffset: number;
    descriptorLength: number;
    centralHeaderOffset: number;
    centralExtraOffset: number;
    centralExtraLength: number;
  }>;
};

describe("ZIP64 readZip", () => {
  test("refuses an undeclared gap before the ZIP64 end record", () => {
    const original=writeZip(new Map([["a.xml",encoder.encode("x")]]),{forceZip64:true});
    const end=original.length-22-20-56;
    const gap=new Uint8Array(original.length+1);gap.set(original.subarray(0,end));gap[end]=0x42;gap.set(original.subarray(end),end+1);
    new DataView(gap.buffer).setBigUint64(gap.length-22-20+8,BigInt(end+1),true);
    expectZipError(()=>readZip(gap),"zip-structure-invalid","central-directory");
  });
  test("reads an empty deflate payload under a nonzero native inflate cap",()=>{
    const archive=buildZip64Archive([{name:'empty.xml',blob:new Uint8Array(),method:8}]);
    expect(readZip(archive.bytes).get('empty.xml')).toEqual(new Uint8Array());
  });
  test("reads ZIP64 members that use per-entry extras and 64-bit data descriptors", () => {
    const archive = buildZip64Archive([
      { name: "[Content_Types].xml", blob: encoder.encode("<Types/>"), dataDescriptor: "64sig" },
      { name: "word/document.xml", blob: encoder.encode("<w:document/>"), dataDescriptor: "64sig" },
    ]);

    const parts = readZip(archive.bytes);
    expect(Array.from(parts.keys())).toEqual(["[Content_Types].xml", "word/document.xml"]);
    expect(decoder.decode(parts.get("[Content_Types].xml")!)).toBe("<Types/>");
    expect(decoder.decode(parts.get("word/document.xml")!)).toBe("<w:document/>");
  });

  test("refuses malformed ZIP64 locators", () => {
    const archive = buildZip64Archive([{ name: "word/document.xml", blob: encoder.encode("x") }]);
    const broken = archive.bytes.slice();
    new DataView(broken.buffer, broken.byteOffset, broken.byteLength).setUint32(archive.locatorOffset, 0, true);

    expectZipError(() => readZip(broken), "zip-structure-invalid", "ZIP64 locator");
  });

  test("refuses ZIP64 central counts that cannot fit the declared central span", () => {
    const archive = buildZip64Archive([{ name: "word/document.xml", blob: encoder.encode("x") }]);
    const broken = archive.bytes.slice();
    const view = new DataView(broken.buffer, broken.byteOffset, broken.byteLength);
    setUint64(view, archive.zip64EocdOffset + 24, 10_001n);
    setUint64(view, archive.zip64EocdOffset + 32, 10_001n);

    expectZipError(
      () => readZip(broken, { maxEntries: 20_000 }),
      "zip-structure-invalid",
      "count exceeds",
    );
  });

  test("refuses ZIP64 offsets beyond Number.MAX_SAFE_INTEGER", () => {
    const archive = buildZip64Archive([{ name: "word/document.xml", blob: encoder.encode("x") }]);
    const broken = archive.bytes.slice();
    const view = new DataView(broken.buffer, broken.byteOffset, broken.byteLength);
    setUint64(view, archive.locatorOffset + 8, MAX_SAFE_BIGINT + 1n);

    expectZipError(() => readZip(broken), "zip-zip64-unsupported", "safe integer");
  });

  test("refuses unsupported ZIP64 extensible data sectors", () => {
    const archive = buildZip64Archive(
      [{ name: "word/document.xml", blob: encoder.encode("x") }],
      { extensibleData: encoder.encode("ext") },
    );

    expectZipError(() => readZip(archive.bytes), "zip-zip64-unsupported", "extensible");
  });

  test("refuses missing central ZIP64 extras for sentinel fields", () => {
    const archive = buildZip64Archive([{ name: "word/document.xml", blob: encoder.encode("x") }]);
    const broken = archive.bytes.slice();
    const view = new DataView(broken.buffer, broken.byteOffset, broken.byteLength);
    view.setUint16(archive.entries[0]!.centralExtraOffset, 0xaaaa, true);

    expectZipError(() => readZip(broken), "zip-structure-invalid", "missing required ZIP64 extra");
  });

  test("refuses duplicate central ZIP64 extras rather than guessing", () => {
    const archive = buildZip64Archive([
      { name: "word/document.xml", blob: encoder.encode("x"), duplicateCentralZip64Extra: true },
    ]);

    expectZipError(() => readZip(archive.bytes), "zip-zip64-unsupported", "duplicate ZIP64 extra");
  });

  test("refuses truncated central ZIP64 extras", () => {
    const archive = buildZip64Archive([{ name: "word/document.xml", blob: encoder.encode("x") }]);
    const broken = archive.bytes.slice();
    const view = new DataView(broken.buffer, broken.byteOffset, broken.byteLength);
    view.setUint16(archive.entries[0]!.centralExtraOffset + 2, 8, true);

    expectZipError(() => readZip(broken), "zip-structure-invalid", "truncated ZIP64 extra");
  });

  test("refuses ZIP64 local and central size mismatches", () => {
    const archive = buildZip64Archive([{ name: "word/document.xml", blob: encoder.encode("payload") }]);
    const broken = archive.bytes.slice();
    const view = new DataView(broken.buffer, broken.byteOffset, broken.byteLength);
    setUint64(view, archive.entries[0]!.localExtraOffset + 4, 99n);

    expectZipError(() => readZip(broken), "zip-local-metadata-mismatch", "local and central");
  });

  test("refuses ZIP64 data descriptor mismatches", () => {
    const archive = buildZip64Archive([
      { name: "word/document.xml", blob: encoder.encode("payload"), dataDescriptor: "64sig" },
    ]);
    const broken = archive.bytes.slice();
    const entry = archive.entries[0]!;
    const view = new DataView(broken.buffer, broken.byteOffset, broken.byteLength);
    setUint64(view, entry.descriptorOffset + 12, 99n);

    expectZipError(() => readZip(broken), "zip-local-metadata-mismatch", "local and central");
  });

  test("refuses CRC mismatches in ZIP64 archives", () => {
    const archive = buildZip64Archive([{ name: "word/document.xml", blob: encoder.encode("payload") }]);
    const broken = archive.bytes.slice();
    const dataOffset=archive.entries[0]!.dataOffset;
    broken[dataOffset] = broken[dataOffset]! ^ 0x01;

    expectZipError(() => readZip(broken), "zip-crc-mismatch", "CRC");
  });
});

describe("ZIP64 writeZip", () => {
  test("writes forced tiny ZIP64 output that readZip can reopen", () => {
    const parts = new Map<string, Uint8Array>([
      ["word/document.xml", encoder.encode("<w:document/>")],
      ["[Content_Types].xml", encoder.encode("<Types/>")],
    ]);

    const archive = writeZip(parts, { forceZip64: true });
    const layout = describeZip64Archive(archive);
    expect(layout.isZip64).toBe(true);
    expect(layout.totalEntries).toBe(2);

    const reopened = readZip(archive);
    expect(decoder.decode(reopened.get("word/document.xml")!)).toBe("<w:document/>");
    expect(decoder.decode(reopened.get("[Content_Types].xml")!)).toBe("<Types/>");
  });

  test("automatically emits ZIP64 EOCD records for 65535 tiny entries", () => {
    const count = 65_535;
    const parts = new Map<string, Uint8Array>();
    for (let index = 0; index < count; index += 1) {
      parts.set(`tiny/${index.toString(36).padStart(4, "0")}.bin`, new Uint8Array());
    }

    const archive = writeZip(parts);
    const layout = describeZip64Archive(archive);
    expect(layout.isZip64).toBe(true);
    expect(layout.totalEntries).toBe(count);

    const reopened = readZip(archive, {
      maxEntries: count,
      maxTotalBytes: count,
      maxArchiveBytes: archive.length + 1,
    });
    expect(reopened.size).toBe(count);
    expect(reopened.get("tiny/0000.bin")).toEqual(new Uint8Array());
    expect(reopened.get(`tiny/${(count - 1).toString(36).padStart(4, "0")}.bin`)).toEqual(new Uint8Array());
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

function buildZip64Archive(
  members: Zip64MemberSpec[],
  options: Zip64BuildOptions = {},
): BuiltZip64Archive {
  const bodyChunks: Uint8Array[] = [];
  const centralChunks: Uint8Array[] = [];
  const entries: BuiltZip64Archive["entries"] = [];
  let bodyLength = 0;

  for (const spec of members) {
    const blob = spec.blob ?? encoder.encode(spec.name);
    const method = spec.method ?? METHOD_STORED;
    const compressed = compress(method, blob);
    const crc = crc32(blob);
    const descriptorKind = spec.dataDescriptor ?? "none";
    const hasDescriptor = descriptorKind !== "none";
    const flags = FLAG_UTF8 | (hasDescriptor ? FLAG_DATA_DESCRIPTOR : 0);
    const localZip64 = spec.includeLocalZip64 ?? true;
    const centralZip64 = spec.includeCentralZip64 ?? true;
    const nameBytes = encoder.encode(spec.name);
    const localExtra = concatBytes(
      localZip64 ? buildZip64Extra({ fileSize: BigInt(blob.length), compressedSize: BigInt(compressed.length) }) : new Uint8Array(),
      spec.duplicateLocalZip64Extra
        ? buildZip64Extra({ fileSize: BigInt(blob.length), compressedSize: BigInt(compressed.length) })
        : new Uint8Array(),
    );
    const localHeaderOffset = bodyLength;
    const localExtraOffset = localHeaderOffset + 30 + nameBytes.length;
    const localHeader = buildLocalHeader({
      versionNeeded: localZip64 ? 45 : 20,
      flags,
      method,
      crc: hasDescriptor ? 0 : crc,
      compressedSize: hasDescriptor ? (localZip64 ? MAX_UINT32 : 0) : (localZip64 ? MAX_UINT32 : compressed.length),
      fileSize: hasDescriptor ? (localZip64 ? MAX_UINT32 : 0) : (localZip64 ? MAX_UINT32 : blob.length),
      nameLength: nameBytes.length,
      extraLength: localExtra.length,
    });
    const descriptor = buildDescriptor({
      kind: descriptorKind,
      crc,
      compressedSize: BigInt(compressed.length),
      fileSize: BigInt(blob.length),
    });
    const dataOffset = localExtraOffset + localExtra.length;
    const descriptorOffset = dataOffset + compressed.length;

    bodyChunks.push(localHeader, nameBytes, localExtra, compressed, descriptor);
    bodyLength += localHeader.length + nameBytes.length + localExtra.length + compressed.length + descriptor.length;

    entries.push({
      name: spec.name,
      localHeaderOffset,
      localExtraOffset,
      localExtraLength: localExtra.length,
      dataOffset,
      dataLength: compressed.length,
      descriptorOffset,
      descriptorLength: descriptor.length,
      centralHeaderOffset: 0,
      centralExtraOffset: 0,
      centralExtraLength: 0,
    });

    const centralExtra = concatBytes(
      centralZip64
        ? buildZip64Extra({
            fileSize: BigInt(blob.length),
            compressedSize: BigInt(compressed.length),
            localHeaderOffset: BigInt(localHeaderOffset),
            diskStart: 0,
          })
        : new Uint8Array(),
      spec.duplicateCentralZip64Extra
        ? buildZip64Extra({
            fileSize: BigInt(blob.length),
            compressedSize: BigInt(compressed.length),
            localHeaderOffset: BigInt(localHeaderOffset),
            diskStart: 0,
          })
        : new Uint8Array(),
    );
    const entry = entries[entries.length - 1]!;
    entry.centralHeaderOffset = 0;
    entry.centralExtraOffset = 0;
    entry.centralExtraLength = centralExtra.length;

    centralChunks.push(
      buildCentralHeader({
        versionMadeBy: centralZip64 ? 45 : 20,
        versionNeeded: centralZip64 ? 45 : 20,
        flags,
        method,
        crc,
        compressedSize: centralZip64 ? MAX_UINT32 : compressed.length,
        fileSize: centralZip64 ? MAX_UINT32 : blob.length,
        nameLength: nameBytes.length,
        extraLength: centralExtra.length,
        commentLength: 0,
        diskStart: centralZip64 ? MAX_UINT16 : 0,
        headerOffset: centralZip64 ? MAX_UINT32 : localHeaderOffset,
      }),
      nameBytes,
      centralExtra,
    );
  }

  const centralOffset = bodyLength;
  let centralLength = 0;
  for (let index = 0; index < entries.length; index += 1) {
    const entry = entries[index]!;
    const spec = members[index]!;
    const nameBytes = encoder.encode(spec.name);
    entry.centralHeaderOffset = centralOffset + centralLength;
    entry.centralExtraOffset = entry.centralHeaderOffset + 46 + nameBytes.length;
    centralLength += 46 + nameBytes.length + entry.centralExtraLength;
  }

  const centralDirectory = concatBytes(...centralChunks);
  const zip64EocdOffset = centralOffset + centralDirectory.length;
  const zip64Eocd = buildZip64EndRecord({
    diskNumber: 0,
    centralDisk: 0,
    diskEntries: BigInt(entries.length),
    totalEntries: BigInt(entries.length),
    centralSize: BigInt(centralDirectory.length),
    centralOffset: BigInt(centralOffset),
    extensibleData: options.extensibleData ?? new Uint8Array(),
  });
  const locatorOffset = zip64EocdOffset + zip64Eocd.length;
  const locator = buildZip64Locator({ zip64EndOffset: BigInt(zip64EocdOffset) });
  const eocdOffset = locatorOffset + locator.length;
  const eocd = buildEndRecord({
    diskNumber: 0,
    centralDisk: 0,
    diskEntries: MAX_UINT16,
    totalEntries: MAX_UINT16,
    centralSize: MAX_UINT32,
    centralOffset: MAX_UINT32,
    commentLength: 0,
  });

  return {
    bytes: concatBytes(...bodyChunks, centralDirectory, zip64Eocd, locator, eocd),
    centralOffset,
    centralSize: centralDirectory.length,
    zip64EocdOffset,
    locatorOffset,
    eocdOffset,
    entries,
  };
}

function describeZip64Archive(bytes: Uint8Array): { isZip64: boolean; totalEntries: number } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOffset = findEndOfCentralDirectory(bytes, view);
  const diskEntries = view.getUint16(eocdOffset + 8, true);
  const totalEntries16 = view.getUint16(eocdOffset + 10, true);
  const centralSize = view.getUint32(eocdOffset + 12, true);
  const centralOffset = view.getUint32(eocdOffset + 16, true);
  const isZip64 =
    diskEntries === MAX_UINT16
    || totalEntries16 === MAX_UINT16
    || centralSize === MAX_UINT32
    || centralOffset === MAX_UINT32;
  if (!isZip64) {
    return { isZip64: false, totalEntries: totalEntries16 };
  }

  const locatorOffset = eocdOffset - 20;
  expect(view.getUint32(locatorOffset, true)).toBe(ZIP64_LOCATOR_SIGNATURE);
  const zip64Offset = Number(getUint64(view, locatorOffset + 8));
  expect(view.getUint32(zip64Offset, true)).toBe(ZIP64_END_SIGNATURE);
  return { isZip64: true, totalEntries: Number(getUint64(view, zip64Offset + 32)) };
}

function buildLocalHeader(fields: {
  versionNeeded: number;
  flags: number;
  method: number;
  crc: number;
  compressedSize: number;
  fileSize: number;
  nameLength: number;
  extraLength: number;
}): Uint8Array {
  const bytes = new Uint8Array(30);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, LOCAL_SIGNATURE, true);
  view.setUint16(4, fields.versionNeeded, true);
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
  versionMadeBy: number;
  versionNeeded: number;
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
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, CENTRAL_SIGNATURE, true);
  view.setUint16(4, fields.versionMadeBy, true);
  view.setUint16(6, fields.versionNeeded, true);
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

function buildZip64Extra(fields: {
  fileSize?: bigint;
  compressedSize?: bigint;
  localHeaderOffset?: bigint;
  diskStart?: number;
}): Uint8Array {
  const payloadSize =
    (fields.fileSize === undefined ? 0 : 8)
    + (fields.compressedSize === undefined ? 0 : 8)
    + (fields.localHeaderOffset === undefined ? 0 : 8)
    + (fields.diskStart === undefined ? 0 : 4);
  const bytes = new Uint8Array(4 + payloadSize);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint16(0, ZIP64_EXTRA_TAG, true);
  view.setUint16(2, payloadSize, true);
  let offset = 4;
  if (fields.fileSize !== undefined) {
    setUint64(view, offset, fields.fileSize);
    offset += 8;
  }
  if (fields.compressedSize !== undefined) {
    setUint64(view, offset, fields.compressedSize);
    offset += 8;
  }
  if (fields.localHeaderOffset !== undefined) {
    setUint64(view, offset, fields.localHeaderOffset);
    offset += 8;
  }
  if (fields.diskStart !== undefined) {
    view.setUint32(offset, fields.diskStart, true);
  }
  return bytes;
}

function buildDescriptor(fields: {
  kind: DescriptorKind;
  crc: number;
  compressedSize: bigint;
  fileSize: bigint;
}): Uint8Array {
  switch (fields.kind) {
    case "none":
      return new Uint8Array();
    case "32":
      return buildDescriptorBody(fields.crc, Number(fields.compressedSize), Number(fields.fileSize), false, false);
    case "32sig":
      return buildDescriptorBody(fields.crc, Number(fields.compressedSize), Number(fields.fileSize), true, false);
    case "64":
      return buildDescriptorBody64(fields.crc, fields.compressedSize, fields.fileSize, false);
    case "64sig":
      return buildDescriptorBody64(fields.crc, fields.compressedSize, fields.fileSize, true);
  }
}

function buildDescriptorBody(
  crc: number,
  compressedSize: number,
  fileSize: number,
  withSignature: boolean,
  useZip64: boolean,
): Uint8Array {
  if (useZip64) throw new Error("unexpected ZIP64 flag for 32-bit descriptor");
  const bytes = new Uint8Array(withSignature ? 16 : 12);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;
  if (withSignature) {
    view.setUint32(0, DATA_DESCRIPTOR_SIGNATURE, true);
    offset = 4;
  }
  view.setUint32(offset, crc >>> 0, true);
  view.setUint32(offset + 4, compressedSize >>> 0, true);
  view.setUint32(offset + 8, fileSize >>> 0, true);
  return bytes;
}

function buildDescriptorBody64(
  crc: number,
  compressedSize: bigint,
  fileSize: bigint,
  withSignature: boolean,
): Uint8Array {
  const bytes = new Uint8Array(withSignature ? 24 : 20);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let offset = 0;
  if (withSignature) {
    view.setUint32(0, DATA_DESCRIPTOR_SIGNATURE, true);
    offset = 4;
  }
  view.setUint32(offset, crc >>> 0, true);
  setUint64(view, offset + 4, compressedSize);
  setUint64(view, offset + 12, fileSize);
  return bytes;
}

function buildZip64EndRecord(fields: {
  diskNumber: number;
  centralDisk: number;
  diskEntries: bigint;
  totalEntries: bigint;
  centralSize: bigint;
  centralOffset: bigint;
  extensibleData: Uint8Array;
}): Uint8Array {
  const bytes = new Uint8Array(56 + fields.extensibleData.length);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, ZIP64_END_SIGNATURE, true);
  setUint64(view, 4, BigInt(44 + fields.extensibleData.length));
  view.setUint16(12, 45, true);
  view.setUint16(14, 45, true);
  view.setUint32(16, fields.diskNumber, true);
  view.setUint32(20, fields.centralDisk, true);
  setUint64(view, 24, fields.diskEntries);
  setUint64(view, 32, fields.totalEntries);
  setUint64(view, 40, fields.centralSize);
  setUint64(view, 48, fields.centralOffset);
  bytes.set(fields.extensibleData, 56);
  return bytes;
}

function buildZip64Locator(fields: { zip64EndOffset: bigint }): Uint8Array {
  const bytes = new Uint8Array(20);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, ZIP64_LOCATOR_SIGNATURE, true);
  view.setUint32(4, 0, true);
  setUint64(view, 8, fields.zip64EndOffset);
  view.setUint32(16, 1, true);
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
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
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

function findEndOfCentralDirectory(bytes: Uint8Array, view: DataView): number {
  for (let offset = bytes.length - 22; offset >= Math.max(0, bytes.length - 22 - 0xffff); offset -= 1) {
    if (view.getUint32(offset, true) !== END_SIGNATURE) {
      continue;
    }
    const commentLength = view.getUint16(offset + 20, true);
    if (offset + 22 + commentLength === bytes.length) {
      return offset;
    }
  }
  throw new Error("EOCD not found");
}

function getUint64(view: DataView, offset: number): bigint {
  return BigInt(view.getUint32(offset + 4, true)) << 32n | BigInt(view.getUint32(offset, true));
}

function setUint64(view: DataView, offset: number, value: bigint): void {
  view.setUint32(offset, Number(value & 0xffffffffn), true);
  view.setUint32(offset + 4, Number((value >> 32n) & 0xffffffffn), true);
}
