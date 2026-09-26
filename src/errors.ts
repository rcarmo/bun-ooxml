/** Stable machine-readable refusal from the OOXML package or editing layer. */
export class OoxmlError extends Error {
  constructor(readonly code: string, message: string) {
    super(message);
    this.name = "OoxmlError";
  }
}
