import { OoxmlError } from "../errors.ts";
import { applyEdits, escapeAttribute, parseXml, type XmlElement } from "../xml/index.ts";
import { OpcPackage } from "./package.ts";

const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const MIME_TOKEN = "[!#$%&'*+.^_`|~0-9A-Za-z-]+";
const MIME_PARAMETER = `${MIME_TOKEN}=(?:${MIME_TOKEN}|\"(?:[^\"\\\\\r\n]|\\\\.)*\")`;
const MIME_PATTERN = new RegExp(`^${MIME_TOKEN}\/${MIME_TOKEN}(?:[ \t]*;[ \t]*${MIME_PARAMETER})*$`);

type DefaultEntry = {
  extension: string;
  type: string;
  node: XmlElement;
};

type OverrideEntry = {
  partName: string;
  type: string;
  node: XmlElement;
};

type ContentTypesModel = {
  xml: string;
  root: XmlElement;
  defaults: Map<string, DefaultEntry>;
  overrides: Map<string, OverrideEntry>;
  childTemplateName: string;
};

export function getContentType(pkg: OpcPackage, name: string): string | undefined {
  assertPartName(name);
  const model = readContentTypes(pkg);
  return model.overrides.get(name)?.type ?? model.defaults.get(defaultKey(name))?.type;
}

export function setPartContentType(pkg: OpcPackage, name: string, type: string): void {
  assertPartName(name);
  assertMimeType(type);

  const model = readContentTypes(pkg);
  const existing = model.overrides.get(name);
  if (existing?.type === type) {
    return;
  }

  if (!existing && model.defaults.get(defaultKey(name))?.type === type) {
    return;
  }

  const nextXml = existing
    ? updateOverrideType(model.xml, existing.node, type)
    : addOverride(model, name, type);

  if (nextXml !== model.xml) {
    pkg.set("[Content_Types].xml", nextXml);
  }
}

export function removePartContentType(pkg: OpcPackage, name: string): void {
  assertPartName(name);

  const model = readContentTypes(pkg);
  const existing = model.overrides.get(name);
  if (!existing) {
    return;
  }

  const nextXml = applyEdits(model.xml, [{
    start: existing.node.start,
    end: existing.node.end,
    value: "",
  }]);

  if (nextXml !== model.xml) {
    pkg.set("[Content_Types].xml", nextXml);
  }
}

function readContentTypes(pkg: OpcPackage): ContentTypesModel {
  const xml = pkg.text("[Content_Types].xml");
  const document = parseXml(xml);
  const root = document.root;
  if (root.localName !== "Types" || root.namespaceURI !== CONTENT_TYPES_NS) {
    throw new OoxmlError("opc-content-types-invalid", "Invalid content types root");
  }

  const defaults = new Map<string, DefaultEntry>();
  const overrides = new Map<string, OverrideEntry>();
  const childTemplateName = chooseChildTemplateName(root);

  for (const child of root.children) {
    const type = child.attributes.ContentType;
    if (child.namespaceURI !== CONTENT_TYPES_NS || !type) {
      throw new OoxmlError("opc-content-types-invalid", "Malformed content type");
    }
    assertMimeType(type, "opc-content-types-invalid");

    if (child.localName === "Default") {
      const extension = child.attributes.Extension;
      if (!extension) {
        throw new OoxmlError("opc-content-types-invalid", "Duplicate/missing default extension");
      }
      const key = extension.toLowerCase();
      if (defaults.has(key)) {
        throw new OoxmlError("opc-content-types-invalid", "Duplicate/missing default extension");
      }
      defaults.set(key, { extension, type, node: child });
      continue;
    }

    if (child.localName === "Override") {
      const partName = child.attributes.PartName;
      if (!partName?.startsWith("/")) {
        throw new OoxmlError("opc-content-types-invalid", "Duplicate/missing part override");
      }
      const canonical = partName.slice(1);
      assertPartName(canonical, "opc-content-types-invalid");
      if (overrides.has(canonical)) {
        throw new OoxmlError("opc-content-types-invalid", "Duplicate/missing part override");
      }
      overrides.set(canonical, { partName: canonical, type, node: child });
      continue;
    }

    throw new OoxmlError("opc-content-types-invalid", "Unknown content type element");
  }

  return { xml, root, defaults, overrides, childTemplateName };
}

function chooseChildTemplateName(root: XmlElement): string {
  const separator = root.name.indexOf(":");
  if (separator !== -1) {
    return `${root.name.slice(0, separator)}:Override`;
  }
  if (root.attributes.xmlns === CONTENT_TYPES_NS) {
    return "Override";
  }
  for (const [name, value] of Object.entries(root.attributes)) {
    if (!name.startsWith("xmlns:") || value !== CONTENT_TYPES_NS) {
      continue;
    }
    return `${name.slice("xmlns:".length)}:Override`;
  }
  throw new OoxmlError("opc-content-types-invalid", "Malformed content types root");
}

function addOverride(model: ContentTypesModel, name: string, type: string): string {
  const overrideXml = `<${model.childTemplateName} PartName="/${escapeAttribute(name)}" ContentType="${escapeAttribute(type)}"/>`;

  if (model.root.selfClosing) {
    const rootTag = model.xml.slice(model.root.start, model.root.openEnd);
    if (!rootTag.endsWith("/>")) {
      throw new OoxmlError("opc-content-types-invalid", "Malformed content types root");
    }
    return applyEdits(model.xml, [{
      start: model.root.start,
      end: model.root.openEnd,
      value: `${rootTag.slice(0, -2)}>${overrideXml}</${model.root.name}>`,
    }]);
  }

  return applyEdits(model.xml, [{
    start: model.root.closeStart,
    end: model.root.closeStart,
    value: overrideXml,
  }]);
}

function updateOverrideType(xml: string, node: XmlElement, type: string): string {
  const range = findAttributeValueRange(xml, node, "ContentType");
  return applyEdits(xml, [{
    start: range.start,
    end: range.end,
    value: escapeAttribute(type),
  }]);
}

function findAttributeValueRange(xml: string, node: XmlElement, attributeName: string): { start: number; end: number } {
  const startTag = xml.slice(node.start, node.openEnd);
  let index = 1;

  while (index < startTag.length && !isTagDelimiter(startTag[index]!)) {
    index += 1;
  }

  while (index < startTag.length) {
    index = skipWhitespace(startTag, index);
    if (index >= startTag.length || startTag[index] === ">" || startTag.startsWith("/>", index)) {
      break;
    }

    const nameStart = index;
    while (index < startTag.length && !isAttributeDelimiter(startTag[index]!)) {
      index += 1;
    }
    const name = startTag.slice(nameStart, index);
    index = skipWhitespace(startTag, index);
    if (startTag[index] !== "=") {
      throw new OoxmlError("opc-content-types-invalid", `Malformed ${node.localName} element`);
    }
    index += 1;
    index = skipWhitespace(startTag, index);

    const quote = startTag[index];
    if (quote !== '"' && quote !== "'") {
      throw new OoxmlError("opc-content-types-invalid", `Malformed ${node.localName} element`);
    }
    index += 1;

    const valueStart = index;
    while (index < startTag.length && startTag[index] !== quote) {
      index += 1;
    }
    if (index >= startTag.length) {
      throw new OoxmlError("opc-content-types-invalid", `Malformed ${node.localName} element`);
    }

    if (name === attributeName) {
      return {
        start: node.start + valueStart,
        end: node.start + index,
      };
    }

    index += 1;
  }

  throw new OoxmlError("opc-content-types-invalid", `Malformed ${node.localName} element`);
}

function skipWhitespace(text: string, index: number): number {
  while (index < text.length && isWhitespace(text[index]!)) {
    index += 1;
  }
  return index;
}

function isWhitespace(char: string): boolean {
  return char === " " || char === "\t" || char === "\n" || char === "\r";
}

function isTagDelimiter(char: string): boolean {
  return isWhitespace(char) || char === "/" || char === ">";
}

function isAttributeDelimiter(char: string): boolean {
  return isWhitespace(char) || char === "=" || char === "/" || char === ">";
}

function defaultKey(name: string): string {
  return name.split(".").at(-1)!.toLowerCase();
}

function assertPartName(name: string, code = "opc-part-name-invalid"): void {
  if (
    !name ||
    name.startsWith("/") ||
    /[%\\\u0000-\u001f?#]/.test(name) ||
    name.split("/").some((segment) => !segment || segment === "." || segment === "..")
  ) {
    throw new OoxmlError(code, code === "opc-part-name-invalid" ? `Noncanonical part name: ${name}` : "Malformed content type");
  }
}

function assertMimeType(type: string, code = "opc-content-type-invalid"): void {
  if (!MIME_PATTERN.test(type)) {
    throw new OoxmlError(code, code === "opc-content-type-invalid" ? `Invalid content type: ${type}` : "Malformed content type");
  }
}
