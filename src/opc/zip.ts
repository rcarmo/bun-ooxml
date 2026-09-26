import { inflateRawSync } from "node:zlib";

import { OoxmlError } from "../errors.ts";

export type ZipLimits = {
  maxEntries?: number;
  maxEntryBytes?: number;
  maxTotalBytes?: number;
  maxCompressionRatio?: number;
  maxArchiveBytes?: number;
};

type NormalizedZipLimits = {
  maxEntries: number;
  maxEntryBytes: number;
  maxTotalBytes: number;
  maxCompressionRatio: number;
  maxArchiveBytes: number;
};

type CentralEntry = {
  index: number;
  name: string;
  nameBytes: Uint8Array;
  flags: number;
  method: number;
  crc: number;
  compressedSize: number;
  fileSize: number;
  localHeaderOffset: number;
  isDirectory: boolean;
  dataStart: number;
  dataEnd: number;
  localEnd: number;
};

type WritableMember = {
  name: string;
  nameBytes: Uint8Array;
  fileSize: number;
  body: Uint8Array;
  method: number;
  crc: number;
};

const LOCAL_SIGNATURE = 0x04034b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;

const FLAG_ENCRYPTED = 0x0001;
const FLAG_DATA_DESCRIPTOR = 0x0008;
const FLAG_UTF8 = 0x0800;

const METHOD_STORED = 0;
const METHOD_DEFLATED = 8;

const LOCAL_FIXED_SIZE = 30;
const CENTRAL_FIXED_SIZE = 46;
const END_FIXED_SIZE = 22;
const MAX_COMMENT_BYTES = 0xffff;
const MAX_UINT16 = 0xffff;
const MAX_UINT32 = 0xffffffff;
const DOS_DATE_1980_01_01 = 33;

const DEFAULT_LIMITS: NormalizedZipLimits = {
  maxEntries: 10_000,
  maxEntryBytes: 128 * 1024 * 1024,
  maxTotalBytes: 512 * 1024 * 1024,
  maxCompressionRatio: 1_000,
  maxArchiveBytes: 256 * 1024 * 1024,
};

const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });
const UTF8_ENCODER = new TextEncoder();

export function readZip(bytes: Uint8Array, limits: ZipLimits = {}): Map<string, Uint8Array> {
  const resolved = normalizeLimits(limits);
  if (bytes.length > resolved.maxArchiveBytes) {
    refuse("zip-archive-too-large", "ZIP archive bytes exceed the configured limit");
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOffset = findEndOfCentralDirectory(bytes, view);

  const diskNumber = view.getUint16(eocdOffset + 4, true);
  const centralDisk = view.getUint16(eocdOffset + 6, true);
  const diskEntries = view.getUint16(eocdOffset + 8, true);
  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const centralSize = view.getUint32(eocdOffset + 12, true);
  const centralOffset = view.getUint32(eocdOffset + 16, true);

  if (diskNumber !== 0 || centralDisk !== 0) {
    refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
  }
  if (
    diskEntries === MAX_UINT16
    || totalEntries === MAX_UINT16
    || centralSize === MAX_UINT32
    || centralOffset === MAX_UINT32
  ) {
    refuse("zip-zip64-unsupported", "ZIP64 metadata is not supported yet");
  }
  if (diskEntries !== totalEntries) {
    refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
  }
  if (totalEntries > resolved.maxEntries) {
    refuse("zip-too-many-entries", "ZIP entry limit exceeded");
  }
  if (centralOffset > eocdOffset || centralOffset + centralSize !== eocdOffset) {
    refuse("zip-structure-invalid", "ZIP central-directory metadata is inconsistent");
  }

  const centralEnd = centralOffset + centralSize;
  const exactNames = new Set<string>();
  const foldedNames = new Map<string, string>();
  const entries: CentralEntry[] = [];

  let cursor = centralOffset;
  for (let index = 0; index < totalEntries; index += 1) {
    if (cursor + CENTRAL_FIXED_SIZE > centralEnd) {
      refuse("zip-structure-invalid", "ZIP central directory is truncated");
    }
    if (view.getUint32(cursor, true) !== CENTRAL_SIGNATURE) {
      refuse("zip-structure-invalid", "ZIP central directory signature is invalid");
    }

    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    const crc = view.getUint32(cursor + 16, true);
    const compressedSize = view.getUint32(cursor + 20, true);
    const fileSize = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const diskStart = view.getUint16(cursor + 34, true);
    const localHeaderOffset = view.getUint32(cursor + 42, true);
    const recordEnd = cursor + CENTRAL_FIXED_SIZE + nameLength + extraLength + commentLength;

    if (recordEnd > centralEnd) {
      refuse("zip-structure-invalid", "ZIP central directory entry overruns its declared size");
    }
    if (diskStart !== 0) {
      refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
    }
    if (localHeaderOffset === MAX_UINT32) {
      refuse("zip-zip64-unsupported", "ZIP64 local offsets are not supported yet");
    }
    ensureSupportedFlags(flags, undefined);
    ensureSupportedMethod(method, undefined);

    const nameBytes = bytes.slice(cursor + CENTRAL_FIXED_SIZE, cursor + CENTRAL_FIXED_SIZE + nameLength);
    const name = decodeZipName(nameBytes);
    const isDirectory = validateMemberName(name);
    ensureDistinctName(name, exactNames, foldedNames);

    entries.push({
      index,
      name,
      nameBytes,
      flags,
      method,
      crc,
      compressedSize,
      fileSize,
      localHeaderOffset,
      isDirectory,
      dataStart: 0,
      dataEnd: 0,
      localEnd: 0,
    });

    cursor = recordEnd;
  }

  if (cursor !== centralEnd) {
    refuse("zip-structure-invalid", "ZIP central directory has undeclared trailing metadata");
  }

  const sortedByOffset = [...entries].sort((left, right) => {
    if (left.localHeaderOffset !== right.localHeaderOffset) {
      return left.localHeaderOffset - right.localHeaderOffset;
    }
    return left.index - right.index;
  });

  let expectedOffset = 0;
  for (let index = 0; index < sortedByOffset.length; index += 1) {
    const entry = sortedByOffset[index]!;
    if (entry.localHeaderOffset !== expectedOffset) {
      refuse("zip-local-layout-invalid", "ZIP local file records must have no gaps or overlaps");
    }

    const recordEnd = sortedByOffset[index + 1]?.localHeaderOffset ?? centralOffset;
    if (recordEnd < entry.localHeaderOffset || recordEnd > centralOffset) {
      refuse("zip-local-layout-invalid", "ZIP local file records must have no gaps or overlaps");
    }

    const local = parseLocalRecord(bytes, view, entry, recordEnd);
    entry.dataStart = local.dataStart;
    entry.dataEnd = local.dataEnd;
    entry.localEnd = local.localEnd;
    expectedOffset = local.localEnd;
  }

  if (expectedOffset !== centralOffset) {
    refuse("zip-local-layout-invalid", "ZIP local file records must end exactly at the central directory");
  }

  const parts = new Map<string, Uint8Array>();
  let totalExpanded = 0;
  for (const entry of entries) {
    if (entry.fileSize > resolved.maxEntryBytes) {
      refuse("zip-entry-too-large", `ZIP member \"${entry.name}\" exceeds the configured entry limit`);
    }
    if (totalExpanded + entry.fileSize > resolved.maxTotalBytes) {
      refuse("zip-total-too-large", "ZIP total expanded size exceeds the configured limit");
    }
    ensureCompressionRatio(entry, resolved.maxCompressionRatio);

    const compressed = bytes.subarray(entry.dataStart, entry.dataEnd);
    if (compressed.length !== entry.compressedSize) {
      refuse("zip-size-mismatch", `ZIP member \"${entry.name}\" declared size does not match stored data`);
    }

    const expanded = expandEntry(entry, compressed);
    if (expanded.length !== entry.fileSize) {
      refuse("zip-size-mismatch", `ZIP member \"${entry.name}\" declared size does not match expanded data`);
    }
    if (crc32(expanded) !== entry.crc) {
      refuse("zip-crc-mismatch", `ZIP member \"${entry.name}\" failed its CRC check`);
    }
    if (entry.isDirectory) {
      if (expanded.length !== 0) {
        refuse("zip-directory-entry-invalid", `ZIP directory entry \"${entry.name}\" must be empty`);
      }
      continue;
    }

    parts.set(entry.name, expanded);
    totalExpanded += expanded.length;
  }

  return parts;
}

export function writeZip(parts: ReadonlyMap<string, Uint8Array>): Uint8Array {
  const exactNames = new Set<string>();
  const foldedNames = new Map<string, string>();
  const members: WritableMember[] = [];

  for (const [name, value] of parts) {
    const isDirectory = validateMemberName(name);
    ensureDistinctName(name, exactNames, foldedNames);

    if (isDirectory) {
      if (value.length !== 0) {
        refuse("zip-directory-entry-invalid", `ZIP directory entry \"${name}\" must be empty`);
      }
      continue;
    }

    const nameBytes = UTF8_ENCODER.encode(name);
    const raw = Uint8Array.from(value);
    const compressed = Uint8Array.from(Bun.deflateSync(raw));
    const useDeflated = compressed.length < raw.length;

    members.push({
      name,
      nameBytes,
      fileSize: raw.length,
      body: useDeflated ? compressed : raw,
      method: useDeflated ? METHOD_DEFLATED : METHOD_STORED,
      crc: crc32(raw),
    });
  }

  members.sort((left, right) => compareBytes(left.nameBytes, right.nameBytes));
  ensureZip32WriteLimits(members);

  const localChunks: Uint8Array[] = [];
  const centralChunks: Uint8Array[] = [];
  let localLength = 0;

  for (const member of members) {
    const headerOffset = localLength;
    const localHeader = createLocalHeader({
      flags: FLAG_UTF8,
      method: member.method,
      crc: member.crc,
      compressedSize: member.body.length,
      fileSize: member.fileSize,
      nameLength: member.nameBytes.length,
      extraLength: 0,
    });

    localChunks.push(localHeader, member.nameBytes, member.body);
    localLength += localHeader.length + member.nameBytes.length + member.body.length;

    const centralHeader = createCentralHeader({
      flags: FLAG_UTF8,
      method: member.method,
      crc: member.crc,
      compressedSize: member.body.length,
      fileSize: member.fileSize,
      nameLength: member.nameBytes.length,
      extraLength: 0,
      commentLength: 0,
      headerOffset,
    });

    centralChunks.push(centralHeader, member.nameBytes);
  }

  const centralDirectory = concatBytes(...centralChunks);
  const centralOffset = localLength;
  const archiveBytes = centralOffset + centralDirectory.length + END_FIXED_SIZE;
  if (archiveBytes > MAX_UINT32) {
    refuse("zip-zip64-unsupported", "ZIP64 output is required for this archive size");
  }

  const endRecord = createEndRecord({
    diskNumber: 0,
    centralDisk: 0,
    diskEntries: members.length,
    totalEntries: members.length,
    centralSize: centralDirectory.length,
    centralOffset,
    commentLength: 0,
  });

  return concatBytes(...localChunks, centralDirectory, endRecord);
}

export function crc32(bytes: Uint8Array): number {
  const native = Bun.hash?.crc32;
  if (typeof native === "function") {
    return native(bytes) >>> 0;
  }

  let value = 0xffffffff;
  for (const byte of bytes) {
    value = CRC32_TABLE[(value ^ byte) & 0xff]! ^ (value >>> 8);
  }
  return (value ^ 0xffffffff) >>> 0;
}

function normalizeLimits(limits: ZipLimits): NormalizedZipLimits {
  return {
    maxEntries: limits.maxEntries ?? DEFAULT_LIMITS.maxEntries,
    maxEntryBytes: limits.maxEntryBytes ?? DEFAULT_LIMITS.maxEntryBytes,
    maxTotalBytes: limits.maxTotalBytes ?? DEFAULT_LIMITS.maxTotalBytes,
    maxCompressionRatio: limits.maxCompressionRatio ?? DEFAULT_LIMITS.maxCompressionRatio,
    maxArchiveBytes: limits.maxArchiveBytes ?? DEFAULT_LIMITS.maxArchiveBytes,
  };
}

function findEndOfCentralDirectory(bytes: Uint8Array, view: DataView): number {
  if (bytes.length < END_FIXED_SIZE) {
    refuse("zip-end-record-missing", "Could not locate ZIP end-of-central-directory record");
  }

  const searchStart = Math.max(0, bytes.length - (END_FIXED_SIZE + MAX_COMMENT_BYTES));
  for (let offset = bytes.length - END_FIXED_SIZE; offset >= searchStart; offset -= 1) {
    if (view.getUint32(offset, true) !== END_SIGNATURE) {
      continue;
    }
    const commentLength = view.getUint16(offset + 20, true);
    if (offset + END_FIXED_SIZE + commentLength === bytes.length) {
      return offset;
    }
  }

  refuse("zip-end-record-missing", "Could not locate ZIP end-of-central-directory record");
}

function ensureSupportedFlags(flags: number, name: string | undefined): void {
  if ((flags & FLAG_ENCRYPTED) !== 0) {
    const label = name === undefined ? "ZIP member" : `ZIP member \"${name}\"`;
    refuse("zip-encryption-unsupported", `${label} is encrypted and cannot be read safely`);
  }
}

function ensureSupportedMethod(method: number, name: string | undefined): void {
  if (method === METHOD_STORED || method === METHOD_DEFLATED) {
    return;
  }

  const label = name === undefined ? "ZIP member" : `ZIP member \"${name}\"`;
  refuse(
    "zip-method-unsupported",
    `${label} uses unsupported compression method ${method}`,
  );
}

function decodeZipName(nameBytes: Uint8Array): string {
  try {
    return UTF8_DECODER.decode(nameBytes);
  } catch {
    refuse("zip-name-invalid", "ZIP member name is not valid UTF-8");
  }
}

function validateMemberName(name: string): boolean {
  if (name.length === 0 || name.startsWith("/") || name.includes("\\")) {
    refuse("zip-name-invalid", `ZIP member name \"${name}\" is noncanonical`);
  }

  const isDirectory = name.endsWith("/");
  const normalized = isDirectory ? name.slice(0, -1) : name;
  if (normalized.length === 0) {
    refuse("zip-name-invalid", `ZIP member name \"${name}\" is noncanonical`);
  }

  const segments = normalized.split("/");
  for (const segment of segments) {
    if (segment.length === 0 || segment === "." || segment === "..") {
      refuse("zip-name-invalid", `ZIP member name \"${name}\" is noncanonical`);
    }
    for (let index = 0; index < segment.length; index += 1) {
      if (segment.charCodeAt(index) < 0x20) {
        refuse("zip-name-invalid", `ZIP member name \"${name}\" is noncanonical`);
      }
    }
  }

  return isDirectory;
}

function ensureDistinctName(name: string, exactNames: Set<string>, foldedNames: Map<string, string>): void {
  if (exactNames.has(name)) {
    refuse("zip-duplicate-entry", `Found duplicate ZIP member \"${name}\"`);
  }

  const folded = asciiFold(name);
  const previous = foldedNames.get(folded);
  if (previous !== undefined && previous !== name) {
    refuse(
      "zip-case-collision",
      `ZIP member name \"${name}\" ASCII case-collides with \"${previous}\"`,
    );
  }

  exactNames.add(name);
  foldedNames.set(folded, name);
}

function asciiFold(value: string): string {
  let folded = "";
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    folded += String.fromCharCode(code >= 0x41 && code <= 0x5a ? code + 0x20 : code);
  }
  return folded;
}

function parseLocalRecord(
  bytes: Uint8Array,
  view: DataView,
  entry: CentralEntry,
  recordEnd: number,
): { dataStart: number; dataEnd: number; localEnd: number } {
  const offset = entry.localHeaderOffset;
  if (offset + LOCAL_FIXED_SIZE > recordEnd) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }
  if (view.getUint32(offset, true) !== LOCAL_SIGNATURE) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }

  const flags = view.getUint16(offset + 6, true);
  const method = view.getUint16(offset + 8, true);
  const crc = view.getUint32(offset + 14, true);
  const compressedSize = view.getUint32(offset + 18, true);
  const fileSize = view.getUint32(offset + 22, true);
  const nameLength = view.getUint16(offset + 26, true);
  const extraLength = view.getUint16(offset + 28, true);
  const headerEnd = offset + LOCAL_FIXED_SIZE + nameLength + extraLength;

  if (headerEnd > recordEnd) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }
  ensureSupportedFlags(flags, entry.name);
  ensureSupportedMethod(method, entry.name);

  if (flags !== entry.flags || method !== entry.method) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }

  const localNameBytes = bytes.subarray(offset + LOCAL_FIXED_SIZE, offset + LOCAL_FIXED_SIZE + nameLength);
  if (!bytesEqual(localNameBytes, entry.nameBytes)) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }

  const dataStart = headerEnd;
  let dataEnd = recordEnd;
  if ((entry.flags & FLAG_DATA_DESCRIPTOR) !== 0) {
    if (
      (crc !== 0 && crc !== entry.crc)
      || (compressedSize !== 0 && compressedSize !== entry.compressedSize)
      || (fileSize !== 0 && fileSize !== entry.fileSize)
    ) {
      refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
    }
    dataEnd = parseDataDescriptor(bytes, view, entry, dataStart, recordEnd);
  } else if (
    crc !== entry.crc
    || compressedSize !== entry.compressedSize
    || fileSize !== entry.fileSize
  ) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }

  return {
    dataStart,
    dataEnd,
    localEnd: recordEnd,
  };
}

function parseDataDescriptor(
  _bytes: Uint8Array,
  view: DataView,
  entry: CentralEntry,
  dataStart: number,
  recordEnd: number,
): number {
  if (recordEnd - dataStart < 12) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }

  let descriptorStart = recordEnd - 12;
  let valueOffset = descriptorStart;
  if (recordEnd - dataStart >= 16 && view.getUint32(recordEnd - 16, true) === DATA_DESCRIPTOR_SIGNATURE) {
    descriptorStart = recordEnd - 16;
    valueOffset = descriptorStart + 4;
  }

  const crc = view.getUint32(valueOffset, true);
  const compressedSize = view.getUint32(valueOffset + 4, true);
  const fileSize = view.getUint32(valueOffset + 8, true);
  if (crc !== entry.crc || compressedSize !== entry.compressedSize || fileSize !== entry.fileSize) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }

  return descriptorStart;
}

function ensureCompressionRatio(entry: CentralEntry, maxCompressionRatio: number): void {
  if (entry.fileSize === 0) {
    return;
  }

  const ratio = entry.compressedSize === 0 ? Number.POSITIVE_INFINITY : entry.fileSize / entry.compressedSize;
  if (ratio > maxCompressionRatio) {
    refuse(
      "zip-compression-ratio-exceeded",
      `ZIP member \"${entry.name}\" exceeds the configured compression ratio limit`,
    );
  }
}

function expandEntry(entry: CentralEntry, compressed: Uint8Array): Uint8Array {
  if (entry.method === METHOD_STORED) {
    return compressed.slice();
  }

  try {
    return new Uint8Array(inflateRawSync(compressed, { maxOutputLength: entry.fileSize }));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes("larger than") || message.includes("maxOutputLength")) {
      refuse("zip-size-mismatch", `ZIP member \"${entry.name}\" declared size does not match expanded data`);
    }
    refuse("zip-data-invalid", `ZIP member \"${entry.name}\" has invalid deflate data`);
  }
}

function ensureZip32WriteLimits(members: WritableMember[]): void {
  if (members.length > MAX_UINT16 - 1) {
    refuse("zip-zip64-unsupported", "ZIP64 output is required for this entry count");
  }

  let bodyLength = 0;
  let centralLength = 0;
  for (const member of members) {
    if (member.nameBytes.length > MAX_UINT16) {
      refuse("zip-zip64-unsupported", `ZIP member \"${member.name}\" name is too long for ZIP32`);
    }
    if (member.fileSize > MAX_UINT32 || member.body.length > MAX_UINT32) {
      refuse("zip-zip64-unsupported", `ZIP member \"${member.name}\" is too large for ZIP32`);
    }

    bodyLength += LOCAL_FIXED_SIZE + member.nameBytes.length + member.body.length;
    centralLength += CENTRAL_FIXED_SIZE + member.nameBytes.length;
    if (bodyLength > MAX_UINT32 || centralLength > MAX_UINT32 || bodyLength + centralLength + END_FIXED_SIZE > MAX_UINT32) {
      refuse("zip-zip64-unsupported", "ZIP64 output is required for this archive size");
    }
  }
}

function createLocalHeader(fields: {
  flags: number;
  method: number;
  crc: number;
  compressedSize: number;
  fileSize: number;
  nameLength: number;
  extraLength: number;
}): Uint8Array {
  const bytes = new Uint8Array(LOCAL_FIXED_SIZE);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, LOCAL_SIGNATURE, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, fields.flags, true);
  view.setUint16(8, fields.method, true);
  view.setUint16(10, 0, true);
  view.setUint16(12, DOS_DATE_1980_01_01, true);
  view.setUint32(14, fields.crc >>> 0, true);
  view.setUint32(18, fields.compressedSize >>> 0, true);
  view.setUint32(22, fields.fileSize >>> 0, true);
  view.setUint16(26, fields.nameLength, true);
  view.setUint16(28, fields.extraLength, true);
  return bytes;
}

function createCentralHeader(fields: {
  flags: number;
  method: number;
  crc: number;
  compressedSize: number;
  fileSize: number;
  nameLength: number;
  extraLength: number;
  commentLength: number;
  headerOffset: number;
}): Uint8Array {
  const bytes = new Uint8Array(CENTRAL_FIXED_SIZE);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, CENTRAL_SIGNATURE, true);
  view.setUint16(4, 20, true);
  view.setUint16(6, 20, true);
  view.setUint16(8, fields.flags, true);
  view.setUint16(10, fields.method, true);
  view.setUint16(12, 0, true);
  view.setUint16(14, DOS_DATE_1980_01_01, true);
  view.setUint32(16, fields.crc >>> 0, true);
  view.setUint32(20, fields.compressedSize >>> 0, true);
  view.setUint32(24, fields.fileSize >>> 0, true);
  view.setUint16(28, fields.nameLength, true);
  view.setUint16(30, fields.extraLength, true);
  view.setUint16(32, fields.commentLength, true);
  view.setUint16(34, 0, true);
  view.setUint16(36, 0, true);
  view.setUint32(38, 0, true);
  view.setUint32(42, fields.headerOffset >>> 0, true);
  return bytes;
}

function createEndRecord(fields: {
  diskNumber: number;
  centralDisk: number;
  diskEntries: number;
  totalEntries: number;
  centralSize: number;
  centralOffset: number;
  commentLength: number;
}): Uint8Array {
  const bytes = new Uint8Array(END_FIXED_SIZE);
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

function compareBytes(left: Uint8Array, right: Uint8Array): number {
  const shared = Math.min(left.length, right.length);
  for (let index = 0; index < shared; index += 1) {
    const delta = left[index]! - right[index]!;
    if (delta !== 0) {
      return delta;
    }
  }
  return left.length - right.length;
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  if (left.length !== right.length) {
    return false;
  }
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return false;
    }
  }
  return true;
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

function refuse(code: string, message: string): never {
  throw new OoxmlError(code, message);
}

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = (value & 1) !== 0 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();
