/** Stable machine-readable refusal from the OOXML package or editing layer. */
/** Portable category for known XML syntax failures; unrelated errors stay unclassified. */
export function classifyXmlParseFailure(error: unknown): "malformed-xml" | undefined {
  return error instanceof OoxmlError && ["XML_MALFORMED", "XML_MISMATCHED_TAG"].includes(error.code) ? "malformed-xml" : undefined;
}

export class OoxmlError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
    this.name = "OoxmlError";
  }
}
