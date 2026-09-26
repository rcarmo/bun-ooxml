// Ordered ZIP32 fixture builder; deliberately permits inconsistent headers.
import { deflateRawSync } from "node:zlib";
const encoder = new TextEncoder();

const LOCAL_SIGNATURE = 0x04034b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;

const FLAG_UTF8 = 0x0800;
const FLAG_DATA_DESCRIPTOR = 0x0008;

const METHOD_STORED = 0;
const METHOD_DEFLATED = 8;

export type ZipMemberSpec = {
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

export type ZipBuildOptions = {
  diskNumber?: number;
  centralDisk?: number;
  diskEntries?: number;
  totalEntries?: number;
  centralSizeDelta?: number;
  archiveComment?: Uint8Array;
};

export function buildZip(members: ZipMemberSpec[], options: ZipBuildOptions = {}): Uint8Array {
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

export function concatBytes(...chunks: Uint8Array[]): Uint8Array {
  const total = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
  const merged = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    merged.set(chunk, offset);
    offset += chunk.length;
  }
  return merged;
}

export function crc32Test(bytes: Uint8Array): number {
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
