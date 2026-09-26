/** Bun-native OOXML entry points. See docs/agents/usage.md for safe edit workflows. */
export { OoxmlError } from "./errors.ts";
export { OpcPackage, type PackageDiff, type Relationship } from "./opc/package.ts";
export { Document, Span } from "./docx/index.ts";
export { Presentation } from "./pptx/index.ts";
export { Workbook } from "./xlsx/index.ts";
export { patchOffice, type PatchRequest, type PatchReceipt, type TargetResult } from "./workflow/index.ts";
