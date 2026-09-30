/** Stable machine-readable refusal from the OOXML package or editing layer. */
/** Portable category for known XML syntax failures; unrelated errors stay unclassified. */
export function classifyXmlParseFailure(error: unknown): "malformed-xml" | undefined {
  return error instanceof OoxmlError && ["XML_MALFORMED", "XML_MISMATCHED_TAG"].includes(error.code) ? "malformed-xml" : undefined;
}

const sharedPackageReasons = new Set([
  'opc-part-name-invalid','opc-target-invalid','opc-content-types-invalid','opc-relationship-target-missing','opc-relationship-duplicate','opc-symlink-destination',
  'zip-duplicate-entry','zip-case-collision','zip-name-invalid','zip-encryption-unsupported','zip-method-unsupported','zip-multi-disk-unsupported','zip-structure-invalid','zip-local-metadata-mismatch','zip-crc-mismatch','zip-size-mismatch','zip-end-record-missing','zip-directory-entry-invalid',
  'zip-archive-too-large','zip-too-many-entries','zip-entry-too-large','zip-total-too-large','zip-compression-ratio-exceeded',
]);
/** Structured validation reason; diagnostic strings and unrelated failures are never classified. */
export function classifyPackageFailure(error: unknown): string | undefined {
  return error instanceof OoxmlError && sharedPackageReasons.has(error.code) ? error.code : undefined;
}

export class OoxmlError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
    this.name = "OoxmlError";
  }
}
