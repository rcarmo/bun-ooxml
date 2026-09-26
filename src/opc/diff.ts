import { OoxmlError } from "../errors.ts";
import { parseXml } from "../xml/index.ts";
import { type OpcPackage, sameBytes } from "./package.ts";

const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";

export interface PartChange {
  name: string;
  kind: "added" | "removed" | "changed";
  beforeSha256?: string;
  afterSha256?: string;
  beforeBytes?: number;
  afterBytes?: number;
  beforeContentType?: string;
  afterContentType?: string;
}

export interface PackageDiffReport {
  added: string[];
  removed: string[];
  changed: string[];
  unchanged: string[];
  parts: PartChange[];
}

type ContentTypeIndex = {
  defaults: Map<string, string>;
  overrides: Map<string, string>;
};

export function diffPackages(before: OpcPackage, after: OpcPackage): PackageDiffReport {
  const beforeTypes = indexContentTypes(before);
  const afterTypes = indexContentTypes(after);
  const names = [...new Set([...before.names(), ...after.names()])].sort();
  const report: PackageDiffReport = { added: [], removed: [], changed: [], unchanged: [], parts: [] };

  for (const name of names) {
    const beforeBytes = before.get(name);
    const afterBytes = after.get(name);
    const beforeContentType = beforeBytes ? resolveContentType(beforeTypes, name) : undefined;
    const afterContentType = afterBytes ? resolveContentType(afterTypes, name) : undefined;

    if (!beforeBytes) {
      report.added.push(name);
      report.parts.push({
        name,
        kind: "added",
        afterSha256: sha256(afterBytes!),
        afterBytes: afterBytes!.length,
        afterContentType,
      });
      continue;
    }

    if (!afterBytes) {
      report.removed.push(name);
      report.parts.push({
        name,
        kind: "removed",
        beforeSha256: sha256(beforeBytes),
        beforeBytes: beforeBytes.length,
        beforeContentType,
      });
      continue;
    }

    if (sameBytes(beforeBytes, afterBytes) && beforeContentType === afterContentType) {
      report.unchanged.push(name);
      continue;
    }

    report.changed.push(name);
    report.parts.push({
      name,
      kind: "changed",
      beforeSha256: sha256(beforeBytes),
      afterSha256: sha256(afterBytes),
      beforeBytes: beforeBytes.length,
      afterBytes: afterBytes.length,
      beforeContentType,
      afterContentType,
    });
  }

  return report;
}

function sha256(bytes: Uint8Array): string {
  return new Bun.CryptoHasher("sha256").update(bytes).digest("hex");
}

function indexContentTypes(pkg: OpcPackage): ContentTypeIndex {
  const xml = parseXml(pkg.text("[Content_Types].xml"));
  if (xml.root.localName !== "Types" || xml.root.namespaceURI !== CONTENT_TYPES_NS) {
    throw new OoxmlError("opc-content-types-invalid", "Invalid content types root");
  }

  const defaults = new Map<string, string>();
  const overrides = new Map<string, string>();
  for (const node of xml.root.children) {
    const type = node.attributes.ContentType;
    if (node.namespaceURI !== CONTENT_TYPES_NS || !type) {
      throw new OoxmlError("opc-content-types-invalid", "Malformed content type");
    }

    if (node.localName === "Default") {
      const extension = node.attributes.Extension?.toLowerCase();
      if (!extension || defaults.has(extension)) {
        throw new OoxmlError("opc-content-types-invalid", "Duplicate/missing default extension");
      }
      defaults.set(extension, type);
      continue;
    }

    if (node.localName === "Override") {
      const name = node.attributes.PartName;
      if (!name?.startsWith("/") || overrides.has(name.slice(1))) {
        throw new OoxmlError("opc-content-types-invalid", "Duplicate/missing part override");
      }
      overrides.set(name.slice(1), type);
      continue;
    }

    throw new OoxmlError("opc-content-types-invalid", "Unknown content type element");
  }

  return { defaults, overrides };
}

function resolveContentType(index: ContentTypeIndex, name: string): string | undefined {
  const override = index.overrides.get(name);
  if (override !== undefined) return override;
  const dot = name.lastIndexOf(".");
  if (dot === -1 || dot === name.length - 1) return undefined;
  return index.defaults.get(name.slice(dot + 1).toLowerCase());
}
