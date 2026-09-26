import { inflateRawSync } from "node:zlib";

import { OoxmlError } from "../errors.ts";

export type ZipLimits = {
  maxEntries?: number;
  maxEntryBytes?: number;
  maxTotalBytes?: number;
  maxCompressionRatio?: number;
  maxArchiveBytes?: number;
};

export type ZipWriteOptions = {
  forceZip64?: boolean;
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

type ResolvedDirectory = {
  totalEntries: number;
  centralSize: number;
  centralOffset: number;
  centralBoundary: number;
};

type Zip64ExtraValues = {
  fileSize?: number;
  compressedSize?: number;
  localHeaderOffset?: number;
  diskStart?: number;
};

type PreparedMember = WritableMember & {
  compressedSize: number;
  headerOffset: number;
  localZip64: boolean;
  centralZip64: boolean;
  localExtra: Uint8Array;
  centralExtra: Uint8Array;
};

type PreparedLayout = {
  members: PreparedMember[];
  localLength: number;
  centralLength: number;
  centralOffset: number;
  archiveZip64: boolean;
  archiveBytes: number;
};

const LOCAL_SIGNATURE = 0x04034b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const END_SIGNATURE = 0x06054b50;
const ZIP64_END_SIGNATURE = 0x06064b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
const DATA_DESCRIPTOR_SIGNATURE = 0x08074b50;
const ZIP64_EXTRA_TAG = 0x0001;

const FLAG_ENCRYPTED = 0x0001;
const FLAG_DATA_DESCRIPTOR = 0x0008;
const FLAG_UTF8 = 0x0800;

const METHOD_STORED = 0;
const METHOD_DEFLATED = 8;

const LOCAL_FIXED_SIZE = 30;
const CENTRAL_FIXED_SIZE = 46;
const END_FIXED_SIZE = 22;
const ZIP64_END_FIXED_SIZE = 56;
const ZIP64_LOCATOR_SIZE = 20;
const MAX_COMMENT_BYTES = 0xffff;
const MAX_UINT16 = 0xffff;
const MAX_UINT32 = 0xffffffff;
const DOS_DATE_1980_01_01 = 33;
const ZIP32_VERSION = 20;
const ZIP64_VERSION = 45;
const ZIP64_END_MIN_SIZE = 44n;
const MAX_SAFE_INTEGER_BIGINT = BigInt(Number.MAX_SAFE_INTEGER);

const DEFAULT_LIMITS: NormalizedZipLimits = {
  maxEntries: 10_000,
  maxEntryBytes: 128 * 1024 * 1024,
  maxTotalBytes: 512 * 1024 * 1024,
  maxCompressionRatio: 1_000,
  maxArchiveBytes: 256 * 1024 * 1024,
};

const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });
const UTF8_ENCODER = new TextEncoder();
const EMPTY_BYTES = new Uint8Array();

export function readZip(bytes: Uint8Array, limits: ZipLimits = {}): Map<string, Uint8Array> {
  const resolved = normalizeLimits(limits);
  if (bytes.length > resolved.maxArchiveBytes) {
    refuse("zip-archive-too-large", "ZIP archive bytes exceed the configured limit");
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const eocdOffset = findEndOfCentralDirectory(bytes, view);
  const directory = resolveDirectory(view, eocdOffset);

  if (directory.totalEntries > resolved.maxEntries) {
    refuse("zip-too-many-entries", "ZIP entry limit exceeded");
  }
  if (directory.centralOffset > directory.centralBoundary) {
    refuse("zip-structure-invalid", "ZIP central-directory metadata is inconsistent");
  }
  if (directory.centralSize !== directory.centralBoundary - directory.centralOffset) {
    refuse("zip-structure-invalid", "ZIP central-directory metadata is inconsistent");
  }
  if (directory.totalEntries > Math.floor(directory.centralSize / CENTRAL_FIXED_SIZE)) {
    refuse("zip-structure-invalid", "ZIP central-directory count exceeds its declared span");
  }

  const centralEnd = directory.centralOffset + directory.centralSize;
  const exactNames = new Set<string>();
  const foldedNames = new Map<string, string>();
  const entries: CentralEntry[] = [];

  let cursor = directory.centralOffset;
  for (let index = 0; index < directory.totalEntries; index += 1) {
    if (cursor + CENTRAL_FIXED_SIZE > centralEnd) {
      refuse("zip-structure-invalid", "ZIP central directory is truncated");
    }
    if (view.getUint32(cursor, true) !== CENTRAL_SIGNATURE) {
      refuse("zip-structure-invalid", "ZIP central directory signature is invalid");
    }

    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    const crc = view.getUint32(cursor + 16, true);
    const rawCompressedSize = view.getUint32(cursor + 20, true);
    const rawFileSize = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const rawDiskStart = view.getUint16(cursor + 34, true);
    const rawLocalHeaderOffset = view.getUint32(cursor + 42, true);
    const recordEnd = cursor + CENTRAL_FIXED_SIZE + nameLength + extraLength + commentLength;

    if (recordEnd > centralEnd) {
      refuse("zip-structure-invalid", "ZIP central directory entry overruns its declared size");
    }
    ensureSupportedFlags(flags, undefined);
    ensureSupportedMethod(method, undefined);

    const nameBytes = bytes.slice(cursor + CENTRAL_FIXED_SIZE, cursor + CENTRAL_FIXED_SIZE + nameLength);
    const name = decodeZipName(nameBytes);
    const isDirectory = validateMemberName(name);
    ensureDistinctName(name, exactNames, foldedNames);

    const extraBytes = bytes.subarray(
      cursor + CENTRAL_FIXED_SIZE + nameLength,
      cursor + CENTRAL_FIXED_SIZE + nameLength + extraLength,
    );
    const zip64 = parseZip64Extra(extraBytes, {
      name,
      location: "central",
      malformedCode: "zip-structure-invalid",
      needsFileSize: rawFileSize === MAX_UINT32,
      needsCompressedSize: rawCompressedSize === MAX_UINT32,
      needsLocalHeaderOffset: rawLocalHeaderOffset === MAX_UINT32,
      needsDiskStart: rawDiskStart === MAX_UINT16,
    });

    const compressedSize = rawCompressedSize === MAX_UINT32 ? requiredZip64Value(zip64.compressedSize, name) : rawCompressedSize;
    const fileSize = rawFileSize === MAX_UINT32 ? requiredZip64Value(zip64.fileSize, name) : rawFileSize;
    const localHeaderOffset =
      rawLocalHeaderOffset === MAX_UINT32 ? requiredZip64Value(zip64.localHeaderOffset, name) : rawLocalHeaderOffset;
    const diskStart = rawDiskStart === MAX_UINT16 ? requiredZip64Value(zip64.diskStart, name) : rawDiskStart;

    if (diskStart !== 0) {
      refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
    }

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

    const recordEnd = sortedByOffset[index + 1]?.localHeaderOffset ?? directory.centralOffset;
    if (recordEnd < entry.localHeaderOffset || recordEnd > directory.centralOffset) {
      refuse("zip-local-layout-invalid", "ZIP local file records must have no gaps or overlaps");
    }

    const local = parseLocalRecord(bytes, view, entry, recordEnd);
    entry.dataStart = local.dataStart;
    entry.dataEnd = local.dataEnd;
    entry.localEnd = local.localEnd;
    expectedOffset = local.localEnd;
  }

  if (expectedOffset !== directory.centralOffset) {
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

export function writeZip(parts: ReadonlyMap<string, Uint8Array>, options: ZipWriteOptions = {}): Uint8Array {
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
    const compressed = raw.length === 0 ? raw : Uint8Array.from(Bun.deflateSync(raw));
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
  const layout = prepareWriteLayout(members, options.forceZip64 ?? false);

  const localChunks: Uint8Array[] = [];
  const centralChunks: Uint8Array[] = [];

  for (const member of layout.members) {
    const localHeader = createLocalHeader({
      versionNeeded: member.localZip64 ? ZIP64_VERSION : ZIP32_VERSION,
      flags: FLAG_UTF8,
      method: member.method,
      crc: member.crc,
      compressedSize: member.localZip64 ? MAX_UINT32 : member.compressedSize,
      fileSize: member.localZip64 ? MAX_UINT32 : member.fileSize,
      nameLength: member.nameBytes.length,
      extraLength: member.localExtra.length,
    });

    localChunks.push(localHeader, member.nameBytes, member.localExtra, member.body);

    const centralHeader = createCentralHeader({
      versionMadeBy: member.centralZip64 ? ZIP64_VERSION : ZIP32_VERSION,
      versionNeeded: member.centralZip64 ? ZIP64_VERSION : ZIP32_VERSION,
      flags: FLAG_UTF8,
      method: member.method,
      crc: member.crc,
      compressedSize: member.centralZip64 ? MAX_UINT32 : member.compressedSize,
      fileSize: member.centralZip64 ? MAX_UINT32 : member.fileSize,
      nameLength: member.nameBytes.length,
      extraLength: member.centralExtra.length,
      commentLength: 0,
      diskStart: member.centralZip64 ? MAX_UINT16 : 0,
      headerOffset: member.centralZip64 ? MAX_UINT32 : member.headerOffset,
    });

    centralChunks.push(centralHeader, member.nameBytes, member.centralExtra);
  }

  const centralDirectory = concatBytes(...centralChunks);
  if (centralDirectory.length !== layout.centralLength) {
    refuse("zip-structure-invalid", "ZIP writer produced inconsistent central-directory bytes");
  }

  if (!layout.archiveZip64) {
    const endRecord = createEndRecord({
      diskNumber: 0,
      centralDisk: 0,
      diskEntries: layout.members.length,
      totalEntries: layout.members.length,
      centralSize: centralDirectory.length,
      centralOffset: layout.centralOffset,
      commentLength: 0,
    });
    return concatBytes(...localChunks, centralDirectory, endRecord);
  }

  const zip64Offset = layout.centralOffset + centralDirectory.length;
  const zip64End = createZip64EndRecord({
    diskNumber: 0,
    centralDisk: 0,
    diskEntries: BigInt(layout.members.length),
    totalEntries: BigInt(layout.members.length),
    centralSize: BigInt(centralDirectory.length),
    centralOffset: BigInt(layout.centralOffset),
  });
  const zip64Locator = createZip64Locator({ zip64EndOffset: BigInt(zip64Offset) });
  const endRecord = createEndRecord({
    diskNumber: 0,
    centralDisk: 0,
    diskEntries: MAX_UINT16,
    totalEntries: MAX_UINT16,
    centralSize: MAX_UINT32,
    centralOffset: MAX_UINT32,
    commentLength: 0,
  });

  return concatBytes(...localChunks, centralDirectory, zip64End, zip64Locator, endRecord);
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

function resolveDirectory(view: DataView, eocdOffset: number): ResolvedDirectory {
  const diskNumber = view.getUint16(eocdOffset + 4, true);
  const centralDisk = view.getUint16(eocdOffset + 6, true);
  const diskEntries = view.getUint16(eocdOffset + 8, true);
  const totalEntries = view.getUint16(eocdOffset + 10, true);
  const centralSize = view.getUint32(eocdOffset + 12, true);
  const centralOffset = view.getUint32(eocdOffset + 16, true);

  if (diskNumber !== 0 || centralDisk !== 0) {
    refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
  }

  const needsZip64 =
    diskEntries === MAX_UINT16
    || totalEntries === MAX_UINT16
    || centralSize === MAX_UINT32
    || centralOffset === MAX_UINT32;
  if (!needsZip64) {
    if (diskEntries !== totalEntries) {
      refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
    }
    return {
      totalEntries,
      centralSize,
      centralOffset,
      centralBoundary: eocdOffset,
    };
  }

  if (eocdOffset < ZIP64_LOCATOR_SIZE) {
    refuse("zip-structure-invalid", "ZIP64 locator is missing or truncated");
  }

  const locatorOffset = eocdOffset - ZIP64_LOCATOR_SIZE;
  if (view.getUint32(locatorOffset, true) !== ZIP64_LOCATOR_SIGNATURE) {
    refuse("zip-structure-invalid", "ZIP64 locator is missing or invalid");
  }

  const locatorDisk = view.getUint32(locatorOffset + 4, true);
  const zip64Offset = toSafeInteger(readUint64(view, locatorOffset + 8), "ZIP64 end-of-central-directory offset");
  const totalDisks = view.getUint32(locatorOffset + 16, true);
  if (locatorDisk !== 0 || totalDisks !== 1) {
    refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
  }
  if (zip64Offset + 12 > locatorOffset) {
    refuse("zip-structure-invalid", "ZIP64 locator offset is out of range");
  }
  if (view.getUint32(zip64Offset, true) !== ZIP64_END_SIGNATURE) {
    refuse("zip-structure-invalid", "ZIP64 end-of-central-directory record is missing or invalid");
  }

  const zip64Size = readUint64(view, zip64Offset + 4);
  if (zip64Size < ZIP64_END_MIN_SIZE) {
    refuse("zip-structure-invalid", "ZIP64 end-of-central-directory record is truncated");
  }
  if (zip64Size !== ZIP64_END_MIN_SIZE) {
    refuse("zip-zip64-unsupported", "ZIP64 extensible data sectors are not supported");
  }

  const zip64End = zip64Offset + 12 + Number(zip64Size);
  if (zip64End !== locatorOffset) {
    refuse("zip-structure-invalid", "ZIP64 locator and end-of-central-directory record are inconsistent");
  }

  const zip64DiskNumber = view.getUint32(zip64Offset + 16, true);
  const zip64CentralDisk = view.getUint32(zip64Offset + 20, true);
  if (zip64DiskNumber !== 0 || zip64CentralDisk !== 0) {
    refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
  }

  const zip64DiskEntries = readUint64(view, zip64Offset + 24);
  const zip64TotalEntries = readUint64(view, zip64Offset + 32);
  const zip64CentralSize = readUint64(view, zip64Offset + 40);
  const zip64CentralOffset = readUint64(view, zip64Offset + 48);

  if (zip64DiskEntries !== zip64TotalEntries) {
    refuse("zip-multi-disk-unsupported", "ZIP multi-disk archives are not supported");
  }

  ensureBackfilledUint16(diskEntries, zip64DiskEntries, "ZIP64 disk entry count");
  ensureBackfilledUint16(totalEntries, zip64TotalEntries, "ZIP64 total entry count");
  ensureBackfilledUint32(centralSize, zip64CentralSize, "ZIP64 central-directory size");
  ensureBackfilledUint32(centralOffset, zip64CentralOffset, "ZIP64 central-directory offset");

  return {
    totalEntries: toSafeInteger(zip64TotalEntries, "ZIP64 total entry count"),
    centralSize: toSafeInteger(zip64CentralSize, "ZIP64 central-directory size"),
    centralOffset: toSafeInteger(zip64CentralOffset, "ZIP64 central-directory offset"),
    centralBoundary: zip64Offset,
  };
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
  const rawCompressedSize = view.getUint32(offset + 18, true);
  const rawFileSize = view.getUint32(offset + 22, true);
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

  const extraBytes = bytes.subarray(offset + LOCAL_FIXED_SIZE + nameLength, headerEnd);
  const zip64 = parseZip64Extra(extraBytes, {
    name: entry.name,
    location: "local",
    malformedCode: "zip-local-metadata-mismatch",
    needsFileSize: rawFileSize === MAX_UINT32,
    needsCompressedSize: rawCompressedSize === MAX_UINT32,
    needsLocalHeaderOffset: false,
    needsDiskStart: false,
  });

  const localCompressedSize =
    rawCompressedSize === MAX_UINT32 ? requiredZip64Value(zip64.compressedSize, entry.name) : rawCompressedSize;
  const localFileSize = rawFileSize === MAX_UINT32 ? requiredZip64Value(zip64.fileSize, entry.name) : rawFileSize;
  const dataStart = headerEnd;
  const dataEnd = dataStart + entry.compressedSize;
  if (dataEnd > recordEnd) {
    refuse("zip-size-mismatch", `ZIP member \"${entry.name}\" declared size overruns its local record`);
  }

  if ((entry.flags & FLAG_DATA_DESCRIPTOR) !== 0) {
    if (!matchesDescriptorPrelude(crc, entry.crc)) {
      refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
    }
    if (!matchesDescriptorSizeField(rawCompressedSize, localCompressedSize, entry.compressedSize)) {
      refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
    }
    if (!matchesDescriptorSizeField(rawFileSize, localFileSize, entry.fileSize)) {
      refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
    }
    parseDataDescriptor(view, entry, dataEnd, recordEnd);
  } else {
    if (
      crc !== entry.crc
      || localCompressedSize !== entry.compressedSize
      || localFileSize !== entry.fileSize
      || dataEnd !== recordEnd
    ) {
      refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
    }
  }

  return {
    dataStart,
    dataEnd,
    localEnd: recordEnd,
  };
}

function parseDataDescriptor(view: DataView, entry: CentralEntry, dataEnd: number, recordEnd: number): void {
  const descriptorLength = recordEnd - dataEnd;
  let valueOffset = dataEnd;
  let zip64 = false;

  if (descriptorLength === 12) {
    valueOffset = dataEnd;
  } else if (descriptorLength === 16) {
    if (view.getUint32(dataEnd, true) !== DATA_DESCRIPTOR_SIGNATURE) {
      refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
    }
    valueOffset = dataEnd + 4;
  } else if (descriptorLength === 20) {
    zip64 = true;
    valueOffset = dataEnd;
  } else if (descriptorLength === 24) {
    if (view.getUint32(dataEnd, true) !== DATA_DESCRIPTOR_SIGNATURE) {
      refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
    }
    zip64 = true;
    valueOffset = dataEnd + 4;
  } else {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }

  const crc = view.getUint32(valueOffset, true);
  if (zip64) {
    const compressedSize = toSafeInteger(
      readUint64(view, valueOffset + 4),
      `ZIP64 data descriptor compressed size for \"${entry.name}\"`,
    );
    const fileSize = toSafeInteger(
      readUint64(view, valueOffset + 12),
      `ZIP64 data descriptor file size for \"${entry.name}\"`,
    );
    if (crc !== entry.crc || compressedSize !== entry.compressedSize || fileSize !== entry.fileSize) {
      refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
    }
    return;
  }

  const compressedSize = view.getUint32(valueOffset + 4, true);
  const fileSize = view.getUint32(valueOffset + 8, true);
  if (crc !== entry.crc || compressedSize !== entry.compressedSize || fileSize !== entry.fileSize) {
    refuse("zip-local-metadata-mismatch", `ZIP local and central metadata disagree for \"${entry.name}\"`);
  }
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
    return new Uint8Array(inflateRawSync(compressed, { maxOutputLength: Math.max(1, entry.fileSize) }));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (message.includes("larger than") || message.includes("maxOutputLength")) {
      refuse("zip-size-mismatch", `ZIP member \"${entry.name}\" declared size does not match expanded data`);
    }
    refuse("zip-data-invalid", `ZIP member \"${entry.name}\" has invalid deflate data`);
  }
}

function prepareWriteLayout(members: WritableMember[], forceZip64: boolean): PreparedLayout {
  const prepared: PreparedMember[] = [];
  let localLength = 0;

  for (const member of members) {
    if (member.nameBytes.length > MAX_UINT16) {
      refuse("zip-zip64-unsupported", `ZIP member \"${member.name}\" name is too long for ZIP output`);
    }

    const localZip64 = forceZip64 || member.fileSize >= MAX_UINT32 || member.body.length >= MAX_UINT32;
    const localExtra = localZip64
      ? createZip64Extra({
          fileSize: BigInt(member.fileSize),
          compressedSize: BigInt(member.body.length),
        })
      : EMPTY_BYTES;
    const headerOffset = localLength;
    localLength = checkedLengthAdd(
      localLength,
      LOCAL_FIXED_SIZE + member.nameBytes.length + localExtra.length + member.body.length,
      `ZIP member \"${member.name}\" local record`,
    );

    prepared.push({
      ...member,
      compressedSize: member.body.length,
      headerOffset,
      localZip64,
      centralZip64: false,
      localExtra,
      centralExtra: EMPTY_BYTES,
    });
  }

  let centralLength = 0;
  for (const member of prepared) {
    member.centralZip64 = forceZip64 || member.localZip64 || member.headerOffset >= MAX_UINT32;
    member.centralExtra = member.centralZip64
      ? createZip64Extra({
          fileSize: BigInt(member.fileSize),
          compressedSize: BigInt(member.compressedSize),
          localHeaderOffset: BigInt(member.headerOffset),
          diskStart: 0,
        })
      : EMPTY_BYTES;
    centralLength = checkedLengthAdd(
      centralLength,
      CENTRAL_FIXED_SIZE + member.nameBytes.length + member.centralExtra.length,
      `ZIP member \"${member.name}\" central record`,
    );
  }

  const centralOffset = localLength;
  const archiveZip64 =
    forceZip64
    || members.length >= MAX_UINT16
    || centralOffset >= MAX_UINT32
    || centralLength >= MAX_UINT32
    || prepared.some((member) => member.centralZip64);
  const archiveBytes = centralOffset + centralLength + (
    archiveZip64 ? ZIP64_END_FIXED_SIZE + ZIP64_LOCATOR_SIZE + END_FIXED_SIZE : END_FIXED_SIZE
  );
  if (!Number.isSafeInteger(archiveBytes)) {
    refuse("zip-zip64-unsupported", "ZIP archive is too large to materialize as a safe integer buffer");
  }

  return {
    members: prepared,
    localLength,
    centralLength,
    centralOffset,
    archiveZip64,
    archiveBytes,
  };
}

function createLocalHeader(fields: {
  versionNeeded: number;
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
  view.setUint16(4, fields.versionNeeded, true);
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
  const bytes = new Uint8Array(CENTRAL_FIXED_SIZE);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, CENTRAL_SIGNATURE, true);
  view.setUint16(4, fields.versionMadeBy, true);
  view.setUint16(6, fields.versionNeeded, true);
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
  view.setUint16(34, fields.diskStart, true);
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

function createZip64EndRecord(fields: {
  diskNumber: number;
  centralDisk: number;
  diskEntries: bigint;
  totalEntries: bigint;
  centralSize: bigint;
  centralOffset: bigint;
}): Uint8Array {
  const bytes = new Uint8Array(ZIP64_END_FIXED_SIZE);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, ZIP64_END_SIGNATURE, true);
  writeUint64(view, 4, ZIP64_END_MIN_SIZE);
  view.setUint16(12, ZIP64_VERSION, true);
  view.setUint16(14, ZIP64_VERSION, true);
  view.setUint32(16, fields.diskNumber, true);
  view.setUint32(20, fields.centralDisk, true);
  writeUint64(view, 24, fields.diskEntries);
  writeUint64(view, 32, fields.totalEntries);
  writeUint64(view, 40, fields.centralSize);
  writeUint64(view, 48, fields.centralOffset);
  return bytes;
}

function createZip64Locator(fields: { zip64EndOffset: bigint }): Uint8Array {
  const bytes = new Uint8Array(ZIP64_LOCATOR_SIZE);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint32(0, ZIP64_LOCATOR_SIGNATURE, true);
  view.setUint32(4, 0, true);
  writeUint64(view, 8, fields.zip64EndOffset);
  view.setUint32(16, 1, true);
  return bytes;
}

function createZip64Extra(fields: {
  fileSize?: bigint;
  compressedSize?: bigint;
  localHeaderOffset?: bigint;
  diskStart?: number;
}): Uint8Array {
  const payloadLength =
    (fields.fileSize === undefined ? 0 : 8)
    + (fields.compressedSize === undefined ? 0 : 8)
    + (fields.localHeaderOffset === undefined ? 0 : 8)
    + (fields.diskStart === undefined ? 0 : 4);
  const bytes = new Uint8Array(4 + payloadLength);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  view.setUint16(0, ZIP64_EXTRA_TAG, true);
  view.setUint16(2, payloadLength, true);
  let offset = 4;
  if (fields.fileSize !== undefined) {
    writeUint64(view, offset, fields.fileSize);
    offset += 8;
  }
  if (fields.compressedSize !== undefined) {
    writeUint64(view, offset, fields.compressedSize);
    offset += 8;
  }
  if (fields.localHeaderOffset !== undefined) {
    writeUint64(view, offset, fields.localHeaderOffset);
    offset += 8;
  }
  if (fields.diskStart !== undefined) {
    view.setUint32(offset, fields.diskStart, true);
  }
  return bytes;
}

function parseZip64Extra(
  extraBytes: Uint8Array,
  options: {
    name: string;
    location: "central" | "local";
    malformedCode: "zip-structure-invalid" | "zip-local-metadata-mismatch";
    needsFileSize: boolean;
    needsCompressedSize: boolean;
    needsLocalHeaderOffset: boolean;
    needsDiskStart: boolean;
  },
): Zip64ExtraValues {
  const label = zip64ExtraLabel(options.location, options.name);
  let cursor = 0;
  let payload: Uint8Array | undefined;
  const extraView = new DataView(extraBytes.buffer, extraBytes.byteOffset, extraBytes.byteLength);

  while (cursor < extraBytes.length) {
    if (cursor + 4 > extraBytes.length) {
      refuse(options.malformedCode, `${label} has a truncated ZIP extra field`);
    }
    const headerId = extraView.getUint16(cursor, true);
    const fieldLength = extraView.getUint16(cursor + 2, true);
    const fieldStart = cursor + 4;
    const fieldEnd = fieldStart + fieldLength;
    if (fieldEnd > extraBytes.length) {
      refuse(options.malformedCode, `${label} has a truncated ZIP64 extra field`);
    }
    if (headerId === ZIP64_EXTRA_TAG) {
      if (payload !== undefined) {
        refuse("zip-zip64-unsupported", `${label} has a duplicate ZIP64 extra field`);
      }
      const required = (options.needsFileSize ? 8 : 0) + (options.needsCompressedSize ? 8 : 0)
        + (options.needsLocalHeaderOffset ? 8 : 0) + (options.needsDiskStart ? 4 : 0);
      if (fieldLength < required) refuse(options.malformedCode, `${label} has a truncated ZIP64 extra field`);
      payload = extraBytes.subarray(fieldStart, fieldEnd);
    }
    cursor = fieldEnd;
  }

  const needsZip64 =
    options.needsFileSize || options.needsCompressedSize || options.needsLocalHeaderOffset || options.needsDiskStart;
  if (!needsZip64) {
    if (payload !== undefined) {
      refuse("zip-zip64-unsupported", `${label} has ambiguous ZIP64 extra data`);
    }
    return {};
  }
  if (payload === undefined) {
    refuse(options.malformedCode, `${label} is missing required ZIP64 extra data`);
  }

  const payloadView = new DataView(payload.buffer, payload.byteOffset, payload.byteLength);
  let offset = 0;
  const resolved: Zip64ExtraValues = {};

  if (options.needsFileSize) {
    if (offset + 8 > payload.length) {
      refuse(options.malformedCode, `${label} has a truncated ZIP64 extra field`);
    }
    resolved.fileSize = toSafeInteger(readUint64(payloadView, offset), `${label} file size`);
    offset += 8;
  }
  if (options.needsCompressedSize) {
    if (offset + 8 > payload.length) {
      refuse(options.malformedCode, `${label} has a truncated ZIP64 extra field`);
    }
    resolved.compressedSize = toSafeInteger(readUint64(payloadView, offset), `${label} compressed size`);
    offset += 8;
  }
  if (options.needsLocalHeaderOffset) {
    if (offset + 8 > payload.length) {
      refuse(options.malformedCode, `${label} has a truncated ZIP64 extra field`);
    }
    resolved.localHeaderOffset = toSafeInteger(readUint64(payloadView, offset), `${label} local header offset`);
    offset += 8;
  }
  if (options.needsDiskStart) {
    if (offset + 4 > payload.length) {
      refuse(options.malformedCode, `${label} has a truncated ZIP64 extra field`);
    }
    resolved.diskStart = payloadView.getUint32(offset, true);
    offset += 4;
  }
  if (offset !== payload.length) {
    refuse("zip-zip64-unsupported", `${label} has ambiguous ZIP64 extra data`);
  }

  return resolved;
}

function zip64ExtraLabel(location: "central" | "local", name: string): string {
  return location === "central"
    ? `ZIP member \"${name}\" central header`
    : `ZIP member \"${name}\" local header`;
}

function readUint64(view: DataView, offset: number): bigint {
  const low = BigInt(view.getUint32(offset, true));
  const high = BigInt(view.getUint32(offset + 4, true));
  return low | (high << 32n);
}

function writeUint64(view: DataView, offset: number, value: bigint): void {
  view.setUint32(offset, Number(value & 0xffffffffn), true);
  view.setUint32(offset + 4, Number((value >> 32n) & 0xffffffffn), true);
}

function toSafeInteger(value: bigint, label: string): number {
  if (value > MAX_SAFE_INTEGER_BIGINT) {
    refuse("zip-zip64-unsupported", `${label} exceeds the safe integer limit`);
  }
  return Number(value);
}

function ensureBackfilledUint16(raw: number, actual: bigint, label: string): void {
  if (raw === MAX_UINT16) {
    return;
  }
  if (actual !== BigInt(raw)) {
    refuse("zip-structure-invalid", `${label} disagrees with the end-of-central-directory record`);
  }
}

function ensureBackfilledUint32(raw: number, actual: bigint, label: string): void {
  if (raw === MAX_UINT32) {
    return;
  }
  if (actual !== BigInt(raw)) {
    refuse("zip-structure-invalid", `${label} disagrees with the end-of-central-directory record`);
  }
}

function requiredZip64Value(value: number | undefined, name: string): number {
  if (value === undefined) {
    refuse("zip-structure-invalid", `ZIP member \"${name}\" is missing required ZIP64 extra data`);
  }
  return value;
}

function matchesDescriptorPrelude(rawCrc: number, expectedCrc: number): boolean {
  return rawCrc === 0 || rawCrc === expectedCrc;
}

function matchesDescriptorSizeField(rawValue: number, resolvedValue: number, expectedValue: number): boolean {
  if (rawValue === 0) {
    return true;
  }
  if (rawValue === MAX_UINT32) {
    return resolvedValue === expectedValue;
  }
  return rawValue === expectedValue;
}

function checkedLengthAdd(total: number, delta: number, label: string): number {
  const next = total + delta;
  if (!Number.isSafeInteger(next)) {
    refuse("zip-zip64-unsupported", `${label} exceeds the safe integer limit`);
  }
  return next;
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
