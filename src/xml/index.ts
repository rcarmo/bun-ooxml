import { OoxmlError } from "../errors.ts";

const XML_NS = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NS = "http://www.w3.org/2000/xmlns/";
const MAX_INPUT_LENGTH = 8 * 1024 * 1024;
const MAX_DEPTH = 256;
const MAX_NODES = 100_000;

export interface XmlDocument {
  elements: XmlElement[];
  root: XmlElement;
}

export interface XmlElement {
  name: string;
  localName: string;
  namespaceURI: string;
  attributes: Record<string, string>;
  children: XmlElement[];
  parent?: XmlElement;
  root: XmlElement;
  text: string;
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  selfClosing: boolean;
}

type ParsedAttribute = {
  name: string;
  value: string;
};

type QualifiedName = {
  prefix: string;
  localName: string;
};

export function parseXml(text: string): XmlDocument {
  if (text.length > MAX_INPUT_LENGTH) {
    fail("XML_INPUT_TOO_LARGE", `XML input exceeds ${MAX_INPUT_LENGTH} UTF-16 code units`);
  }

  validateXmlStringChars(text);

  let index = 0;
  let nodeCount = 0;
  let root: XmlElement | undefined;
  const preorder: XmlElement[] = [];
  const stack: XmlElement[] = [];
  const textBuffers = new Map<XmlElement, string[]>();
  const namespaceStack: Map<string, string>[] = [new Map([["xml", XML_NS]])];

  const readName = (context: string): string => {
    const { name, next } = readXmlName(text, index, context);
    index = next;
    return name;
  };

  const consumeWhitespace = (): number => {
    const start = index;
    while (index < text.length && isWhitespaceChar(text[index])) {
      index += 1;
    }
    return index - start;
  };

  const skipWhitespace = (): void => {
    consumeWhitespace();
  };

  const appendText = (value: string): void => {
    if (value.length === 0) {
      return;
    }

    for (const element of stack) {
      let buffer = textBuffers.get(element);
      if (!buffer) {
        buffer = [];
        textBuffers.set(element, buffer);
      }
      buffer.push(value);
    }
  };

  if (hasXmlDeclarationStart(text)) {
    parseXmlDeclaration();
  }

  while (index < text.length) {
    if (text[index] !== "<") {
      parseText();
      continue;
    }

    if (text.startsWith("<!--", index)) {
      parseComment();
      continue;
    }

    if (text.startsWith("<![CDATA[", index)) {
      parseCdata();
      continue;
    }

    if (text.startsWith("</", index)) {
      parseEndTag();
      continue;
    }

    if (text.startsWith("<?", index)) {
      parseProcessingInstruction();
      continue;
    }

    if (text.startsWith("<!", index)) {
      parseMarkupDeclaration();
      continue;
    }

    parseStartTag();
  }

  if (stack.length > 0) {
    fail("XML_MALFORMED", `Unclosed element <${stack[stack.length - 1]?.name}>`);
  }

  if (!root) {
    fail("XML_MALFORMED", "XML document is missing a root element");
  }

  for (const element of preorder) {
    const segments = textBuffers.get(element);
    if (segments) {
      element.text = segments.join("");
    }
  }

  return { elements: preorder, root };

  function parseXmlDeclaration(): void {
    index += 5;
    if (consumeWhitespace() === 0) {
      fail("XML_MALFORMED", "Malformed XML declaration");
    }

    const version = parseDeclarationPseudoAttribute("version");
    if (!/^1\.[0-9]+$/.test(version)) {
      fail("XML_MALFORMED", "Malformed XML declaration");
    }

    let seenEncoding = false;
    let seenStandalone = false;

    while (true) {
      if (text.startsWith("?>", index)) {
        index += 2;
        return;
      }

      const whitespace = consumeWhitespace();
      if (text.startsWith("?>", index)) {
        index += 2;
        return;
      }
      if (whitespace === 0) {
        fail("XML_MALFORMED", "Malformed XML declaration");
      }

      const name = readName("XML declaration pseudo-attribute");
      if (name === "encoding") {
        if (seenEncoding || seenStandalone) {
          fail("XML_MALFORMED", "Malformed XML declaration");
        }
        seenEncoding = true;
        const encoding = parseDeclarationValue();
        if (!/^[A-Za-z][A-Za-z0-9._-]*$/.test(encoding)) {
          fail("XML_MALFORMED", "Malformed XML declaration");
        }
        continue;
      }

      if (name === "standalone") {
        if (seenStandalone) {
          fail("XML_MALFORMED", "Malformed XML declaration");
        }
        seenStandalone = true;
        const standalone = parseDeclarationValue();
        if (standalone !== "yes" && standalone !== "no") {
          fail("XML_MALFORMED", "Malformed XML declaration");
        }
        continue;
      }

      fail("XML_MALFORMED", "Malformed XML declaration");
    }
  }

  function parseDeclarationPseudoAttribute(expectedName: string): string {
    const name = readName("XML declaration pseudo-attribute");
    if (name !== expectedName) {
      fail("XML_MALFORMED", "Malformed XML declaration");
    }
    return parseDeclarationValue();
  }

  function parseDeclarationValue(): string {
    skipWhitespace();
    if (text[index] !== "=") {
      fail("XML_MALFORMED", "Malformed XML declaration");
    }
    index += 1;
    skipWhitespace();

    const quote = text[index];
    if (quote !== '"' && quote !== "'") {
      fail("XML_MALFORMED", "Malformed XML declaration");
    }
    index += 1;

    const valueStart = index;
    while (index < text.length && text[index] !== quote) {
      const char = text[index];
      if (char === "<" || char === "&") {
        fail("XML_MALFORMED", "Malformed XML declaration");
      }
      index += 1;
    }

    if (index >= text.length) {
      fail("XML_MALFORMED", "Malformed XML declaration");
    }

    const value = text.slice(valueStart, index);
    index += 1;
    return value;
  }

  function parseText(): void {
    const start = index;
    while (index < text.length && text[index] !== "<") {
      index += 1;
    }

    const raw = text.slice(start, index);
    if (raw.includes("]]>")) {
      fail("XML_MALFORMED", "]]> is only allowed inside CDATA sections");
    }

    const decoded = decodeEntities(raw);
    if (stack.length === 0) {
      if (!isXmlWhitespace(decoded)) {
        fail("XML_MALFORMED", "Character data is only allowed inside the root element");
      }
      return;
    }

    appendText(decoded);
  }

  function parseComment(): void {
    const end = text.indexOf("-->", index + 4);
    if (end === -1) {
      fail("XML_MALFORMED", "Unterminated XML comment");
    }

    const content = text.slice(index + 4, end);
    if (content.includes("--") || content.endsWith("-")) {
      fail("XML_MALFORMED", "Malformed XML comment");
    }

    index = end + 3;
  }

  function parseCdata(): void {
    if (stack.length === 0) {
      fail("XML_MALFORMED", "CDATA is only allowed inside the root element");
    }

    const end = text.indexOf("]]>", index + 9);
    if (end === -1) {
      fail("XML_MALFORMED", "Unterminated CDATA section");
    }

    appendText(text.slice(index + 9, end));
    index = end + 3;
  }

  function parseProcessingInstruction(): void {
    index += 2;
    const target = readName("processing instruction");
    if (/^xml$/i.test(target)) {
      fail("XML_MALFORMED", "Processing instruction target 'xml' is reserved");
    }

    if (text.startsWith("?>", index)) {
      index += 2;
      return;
    }

    if (!isWhitespaceChar(text[index])) {
      fail("XML_MALFORMED", `Malformed processing instruction <?${target}>`);
    }
    skipWhitespace();

    const end = text.indexOf("?>", index);
    if (end === -1) {
      fail("XML_MALFORMED", `Unterminated processing instruction <?${target}>`);
    }

    index = end + 2;
  }

  function parseMarkupDeclaration(): void {
    if (text.slice(index, index + 9).toUpperCase() === "<!DOCTYPE") {
      fail("XML_DTD_FORBIDDEN", "DTD declarations are not allowed");
    }

    fail("XML_MALFORMED", "Unsupported markup declaration");
  }

  function parseStartTag(): void {
    const start = index;
    index += 1;
    const name = readName("element");

    if (stack.length + 1 > MAX_DEPTH) {
      fail("XML_DEPTH_LIMIT", `XML nesting depth exceeds ${MAX_DEPTH}`);
    }

    const rawAttributes = new Set<string>();
    const parsedAttributes: ParsedAttribute[] = [];
    let selfClosing = false;
    let terminated = false;

    while (index < text.length) {
      const whitespace = consumeWhitespace();

      if (text.startsWith("/>", index)) {
        selfClosing = true;
        index += 2;
        terminated = true;
        break;
      }

      if (text[index] === ">") {
        index += 1;
        terminated = true;
        break;
      }

      if (whitespace === 0) {
        fail("XML_MALFORMED", `Expected whitespace before attribute in <${name}>`);
      }

      const attributeName = readName("attribute");
      if (rawAttributes.has(attributeName)) {
        fail("XML_DUPLICATE_ATTRIBUTE", `Duplicate attribute '${attributeName}'`);
      }
      rawAttributes.add(attributeName);

      skipWhitespace();
      if (text[index] !== "=") {
        fail("XML_MALFORMED", `Expected '=' after attribute '${attributeName}'`);
      }
      index += 1;
      skipWhitespace();

      const quote = text[index];
      if (quote !== '"' && quote !== "'") {
        fail("XML_MALFORMED", `Attribute '${attributeName}' must use quotes`);
      }
      index += 1;

      const valueStart = index;
      while (index < text.length && text[index] !== quote) {
        if (text[index] === "<") {
          fail("XML_MALFORMED", `Attribute '${attributeName}' cannot contain '<'`);
        }
        index += 1;
      }

      if (index >= text.length) {
        fail("XML_MALFORMED", `Unterminated attribute '${attributeName}'`);
      }

      parsedAttributes.push({
        name: attributeName,
        value: decodeEntities(text.slice(valueStart, index)),
      });
      index += 1;
    }

    if (!terminated) {
      fail("XML_MALFORMED", `Unterminated start tag <${name}>`);
    }

    const openEnd = index;
    const frame = new Map(namespaceStack[namespaceStack.length - 1]);
    for (const attribute of parsedAttributes) {
      applyNamespaceDeclaration(frame, attribute);
    }

    const resolvedName = resolveElementName(name, frame);
    const attributes = Object.create(null) as Record<string, string>;
    const expandedAttributes = new Set<string>();

    for (const attribute of parsedAttributes) {
      attributes[attribute.name] = attribute.value;
      if (isNamespaceDeclaration(attribute.name)) {
        continue;
      }

      const resolvedAttribute = resolveAttributeName(attribute.name, frame);
      const expandedName = `${resolvedAttribute.namespaceURI}\u0000${resolvedAttribute.localName}`;
      if (expandedAttributes.has(expandedName)) {
        fail(
          "XML_DUPLICATE_ATTRIBUTE",
          `Duplicate expanded attribute '${resolvedAttribute.localName}'`,
        );
      }
      expandedAttributes.add(expandedName);
    }

    if (stack.length === 0 && root) {
      fail("XML_MALFORMED", "XML documents may only contain one root element");
    }

    nodeCount += 1;
    if (nodeCount > MAX_NODES) {
      fail("XML_NODE_LIMIT", `XML node count exceeds ${MAX_NODES}`);
    }

    const parent = stack[stack.length - 1];
    const element: XmlElement = {
      name,
      localName: resolvedName.localName,
      namespaceURI: resolvedName.namespaceURI,
      attributes,
      children: [],
      parent,
      root: undefined as unknown as XmlElement,
      text: "",
      start,
      openEnd,
      closeStart: openEnd,
      end: openEnd,
      selfClosing,
    };

    element.root = parent ? parent.root : element;
    parent?.children.push(element);
    preorder.push(element);
    if (!root) {
      root = element;
    }

    if (!selfClosing) {
      stack.push(element);
      namespaceStack.push(frame);
    }
  }

  function parseEndTag(): void {
    if (stack.length === 0) {
      fail("XML_MISMATCHED_TAG", "Closing tag without a matching opening tag");
    }

    const closeStart = index;
    index += 2;
    const name = readName("closing tag");
    skipWhitespace();
    if (text[index] !== ">") {
      fail("XML_MALFORMED", `Malformed closing tag </${name}>`);
    }
    index += 1;

    const current = stack.pop();
    namespaceStack.pop();
    if (!current || current.name !== name) {
      fail("XML_MISMATCHED_TAG", `Expected </${current?.name ?? "?"}> but found </${name}>`);
    }

    current.closeStart = closeStart;
    current.end = index;
  }
}

export function elements(
  xml: XmlDocument | XmlElement,
  localName: string,
  namespaceURI?: string,
): XmlElement[] {
  const nodes = isDocument(xml) ? xml.elements : traverseSubtree(xml);
  return nodes.filter(
    (element) =>
      element.localName === localName &&
      (namespaceURI === undefined || element.namespaceURI === namespaceURI),
  );
}

export function escapeText(value: string): string {
  validateXmlStringChars(value);
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

export function escapeAttribute(value: string): string {
  validateXmlStringChars(value);
  return escapeText(value)
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

export function applyEdits(
  xml: string,
  edits: { start: number; end: number; value: string }[],
): string {
  if (edits.length === 0) {
    return xml;
  }

  const sorted = [...edits].sort((left, right) => left.start - right.start || left.end - right.end);
  let previousEnd = -1;

  for (const edit of sorted) {
    if (
      !Number.isInteger(edit.start) ||
      !Number.isInteger(edit.end) ||
      edit.start < 0 ||
      edit.end < edit.start ||
      edit.end > xml.length
    ) {
      fail("XML_EDIT_BOUNDS", "Edit offsets are out of bounds");
    }

    if (edit.start < previousEnd) {
      fail("XML_EDIT_OVERLAP", "Edit ranges must be disjoint");
    }

    validateXmlStringChars(edit.value);
    previousEnd = edit.end;
  }

  const parts: string[] = [];
  let cursor = 0;
  for (const edit of sorted) {
    if (cursor < edit.start) {
      parts.push(xml.slice(cursor, edit.start));
    }
    parts.push(edit.value);
    cursor = edit.end;
  }
  if (cursor < xml.length) {
    parts.push(xml.slice(cursor));
  }

  const result = parts.join("");

  try {
    parseXml(result);
  } catch (error) {
    throw new OoxmlError(
      "XML_EDIT_UNSAFE",
      `Applying edits produced unsafe or malformed XML: ${error instanceof Error ? error.message : String(error)}`,
    );
  }

  return result;
}

function isDocument(value: XmlDocument | XmlElement): value is XmlDocument {
  return "elements" in value;
}

function traverseSubtree(root: XmlElement): XmlElement[] {
  const result: XmlElement[] = [];
  const stack = [root];
  while (stack.length > 0) {
    const current = stack.pop();
    if (!current) {
      continue;
    }

    result.push(current);
    for (let index = current.children.length - 1; index >= 0; index -= 1) {
      const child = current.children[index];
      if (child) {
        stack.push(child);
      }
    }
  }
  return result;
}

function fail(code: string, message: string): never {
  throw new OoxmlError(code, message);
}

function hasXmlDeclarationStart(text: string): boolean {
  return text.startsWith("<?xml") && isWhitespaceChar(text[5]);
}

function decodeEntities(raw: string): string {
  if (!raw.includes("&")) {
    return raw;
  }

  const parts: string[] = [];
  let chunkStart = 0;
  let index = 0;

  while (index < raw.length) {
    if (raw[index] !== "&") {
      index += 1;
      continue;
    }

    if (index > chunkStart) {
      parts.push(raw.slice(chunkStart, index));
    }

    const semi = raw.indexOf(";", index + 1);
    if (semi === -1) {
      fail("XML_ENTITY_FORBIDDEN", "Unterminated entity reference");
    }

    const entity = raw.slice(index + 1, semi);
    parts.push(decodeEntity(entity));
    index = semi + 1;
    chunkStart = index;
  }

  if (chunkStart < raw.length) {
    parts.push(raw.slice(chunkStart));
  }

  return parts.join("");
}

function decodeEntity(entity: string): string {
  switch (entity) {
    case "amp":
      return "&";
    case "lt":
      return "<";
    case "gt":
      return ">";
    case "apos":
      return "'";
    case "quot":
      return '"';
  }

  if (entity.startsWith("#x") || entity.startsWith("#X")) {
    const digits = entity.slice(2);
    if (!digits || /[^0-9a-fA-F]/.test(digits)) {
      fail("XML_ENTITY_FORBIDDEN", `Unsupported entity '&${entity};'`);
    }
    return codePointToString(Number.parseInt(digits, 16));
  }

  if (entity.startsWith("#")) {
    const digits = entity.slice(1);
    if (!digits || /[^0-9]/.test(digits)) {
      fail("XML_ENTITY_FORBIDDEN", `Unsupported entity '&${entity};'`);
    }
    return codePointToString(Number.parseInt(digits, 10));
  }

  fail("XML_ENTITY_FORBIDDEN", `Unsupported entity '&${entity};'`);
}

function codePointToString(codePoint: number): string {
  if (!Number.isInteger(codePoint) || !isValidXmlCodePoint(codePoint)) {
    fail("XML_INVALID_CHAR", `Invalid XML character reference U+${codePoint.toString(16).toUpperCase()}`);
  }

  return String.fromCodePoint(codePoint);
}

function applyNamespaceDeclaration(frame: Map<string, string>, attribute: ParsedAttribute): void {
  if (attribute.name === "xmlns") {
    if (attribute.value === XML_NS || attribute.value === XMLNS_NS) {
      fail("XML_MALFORMED", "The default namespace URI is reserved");
    }
    frame.set("", attribute.value);
    return;
  }

  const qualified = splitQualifiedName(attribute.name);
  if (qualified.prefix !== "xmlns") {
    return;
  }

  const declaredPrefix = qualified.localName;
  if (declaredPrefix === "xmlns") {
    fail("XML_MALFORMED", "The xmlns prefix is reserved");
  }
  if (declaredPrefix === "xml") {
    if (attribute.value !== XML_NS) {
      fail("XML_MALFORMED", "The xml prefix must remain bound to the XML namespace");
    }
    frame.set("xml", XML_NS);
    return;
  }
  if (attribute.value === "") {
    fail("XML_UNBOUND_PREFIX", `Cannot undeclare namespace prefix '${declaredPrefix}'`);
  }
  if (attribute.value === XML_NS || attribute.value === XMLNS_NS) {
    fail("XML_MALFORMED", `Namespace URI '${attribute.value}' is reserved`);
  }

  frame.set(declaredPrefix, attribute.value);
}

function resolveElementName(name: string, frame: Map<string, string>): QualifiedName & { namespaceURI: string } {
  const qualified = splitQualifiedName(name);
  validateReservedQualifiedName(qualified, "element");
  if (qualified.prefix === "") {
    return {
      ...qualified,
      namespaceURI: frame.get("") ?? "",
    };
  }

  const namespaceURI = frame.get(qualified.prefix);
  if (namespaceURI === undefined) {
    fail("XML_UNBOUND_PREFIX", `Unbound namespace prefix '${qualified.prefix}'`);
  }

  return { ...qualified, namespaceURI };
}

function resolveAttributeName(name: string, frame: Map<string, string>): QualifiedName & { namespaceURI: string } {
  const qualified = splitQualifiedName(name);
  validateReservedQualifiedName(qualified, "attribute");
  if (qualified.prefix === "") {
    return { ...qualified, namespaceURI: "" };
  }
  if (qualified.prefix === "xmlns") {
    return { ...qualified, namespaceURI: XMLNS_NS };
  }

  const namespaceURI = frame.get(qualified.prefix);
  if (namespaceURI === undefined) {
    fail("XML_UNBOUND_PREFIX", `Unbound namespace prefix '${qualified.prefix}'`);
  }

  return { ...qualified, namespaceURI };
}

function isNamespaceDeclaration(name: string): boolean {
  return name === "xmlns" || name.startsWith("xmlns:");
}

function splitQualifiedName(name: string): QualifiedName {
  const first = name.indexOf(":");
  if (first === -1) {
    return { prefix: "", localName: name };
  }

  if (first === 0 || first === name.length - 1 || name.indexOf(":", first + 1) !== -1) {
    fail("XML_MALFORMED", `Malformed qualified name '${name}'`);
  }

  return {
    prefix: name.slice(0, first),
    localName: name.slice(first + 1),
  };
}

function validateReservedQualifiedName(
  qualified: QualifiedName,
  kind: "element" | "attribute",
): void {
  if (qualified.prefix === "xmlns") {
    fail("XML_MALFORMED", `The xmlns prefix is reserved and cannot be used for ${kind} names`);
  }
}

function isWhitespaceChar(value: string | undefined): boolean {
  return value === " " || value === "\t" || value === "\n" || value === "\r";
}

function isXmlWhitespace(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    if (!isWhitespaceChar(value[index])) {
      return false;
    }
  }
  return true;
}

function validateXmlStringChars(value: string): void {
  for (let index = 0; index < value.length; ) {
    const codePoint = value.codePointAt(index);
    if (codePoint === undefined) {
      break;
    }

    if (!isValidXmlCodePoint(codePoint)) {
      fail("XML_INVALID_CHAR", `Invalid XML character U+${codePoint.toString(16).toUpperCase()}`);
    }

    index += codePoint > 0xffff ? 2 : 1;
  }
}

function isValidXmlCodePoint(codePoint: number): boolean {
  return (
    codePoint === 0x9 ||
    codePoint === 0xa ||
    codePoint === 0xd ||
    (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
    (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
    (codePoint >= 0x10000 && codePoint <= 0x10ffff)
  );
}

function isNameStartCodePoint(codePoint: number): boolean {
  return (
    codePoint === 0x3a ||
    codePoint === 0x5f ||
    (codePoint >= 0x41 && codePoint <= 0x5a) ||
    (codePoint >= 0x61 && codePoint <= 0x7a) ||
    (codePoint >= 0xc0 && codePoint <= 0xd6) ||
    (codePoint >= 0xd8 && codePoint <= 0xf6) ||
    (codePoint >= 0xf8 && codePoint <= 0x2ff) ||
    (codePoint >= 0x370 && codePoint <= 0x37d) ||
    (codePoint >= 0x37f && codePoint <= 0x1fff) ||
    (codePoint >= 0x200c && codePoint <= 0x200d) ||
    (codePoint >= 0x2070 && codePoint <= 0x218f) ||
    (codePoint >= 0x2c00 && codePoint <= 0x2fef) ||
    (codePoint >= 0x3001 && codePoint <= 0xd7ff) ||
    (codePoint >= 0xf900 && codePoint <= 0xfdcf) ||
    (codePoint >= 0xfdf0 && codePoint <= 0xfffd) ||
    (codePoint >= 0x10000 && codePoint <= 0xeffff)
  );
}

function isNameCharCodePoint(codePoint: number): boolean {
  return (
    isNameStartCodePoint(codePoint) ||
    codePoint === 0x2d ||
    codePoint === 0x2e ||
    codePoint === 0xb7 ||
    (codePoint >= 0x30 && codePoint <= 0x39) ||
    (codePoint >= 0x300 && codePoint <= 0x36f) ||
    (codePoint >= 0x203f && codePoint <= 0x2040)
  );
}

function readXmlName(text: string, offset: number, context: string): { name: string; next: number } {
  const first = text.codePointAt(offset);
  if (first === undefined || !isNameStartCodePoint(first)) {
    fail("XML_MALFORMED", `Expected ${context} name at offset ${offset}`);
  }

  let next = offset + (first > 0xffff ? 2 : 1);
  while (next < text.length) {
    const codePoint = text.codePointAt(next);
    if (codePoint === undefined || !isNameCharCodePoint(codePoint)) {
      break;
    }
    next += codePoint > 0xffff ? 2 : 1;
  }

  return { name: text.slice(offset, next), next };
}
