import { OoxmlError } from '../errors.ts';
import { escapeAttribute, escapeText, inspectXmlEvents, validateXml } from './index.ts';
import type { AttributeSource } from './attributes.ts';
import type { XmlRemovalTarget } from './removal.ts';

export interface XmlExpandedName { readonly localName: string; readonly namespaceURI: string }
export interface XmlStructuredAttribute { readonly name: XmlExpandedName; readonly value: string }
export type XmlContent = string | { readonly name: XmlExpandedName; readonly attributes?: readonly XmlStructuredAttribute[]; readonly children?: readonly XmlContent[] };
export interface XmlStructurePatch { readonly target: XmlRemovalTarget; readonly children: readonly XmlContent[] }
export interface StructureSource extends AttributeSource {
  name: string; end: number; closeStart: number; depth: number; parent?: StructureSource;
}
const MAX_SIZE = 8 * 1024 * 1024, MAX_ITEMS = 100_000, MAX_DEPTH = 256;
const XML = 'http://www.w3.org/XML/1998/namespace', XMLNS = 'http://www.w3.org/2000/xmlns/';
function fail(code: string, message: string): never { throw new OoxmlError(code, message); }
function arrayLength(value: unknown): number {
  if (!Array.isArray(value)) fail('XML_STRUCTURE_INPUT', 'Expected a structured content array');
  const length = value.length;
  if (!Number.isSafeInteger(length) || length < 0 || length > MAX_ITEMS) fail('XML_STRUCTURE_LIMIT', 'Structured array exceeds limit');
  return length;
}

export function editStructure(source: string, patches: readonly XmlStructurePatch[], owned: WeakMap<XmlRemovalTarget, StructureSource>, replace: boolean): string {
  const count = arrayLength(patches);
  const selected: { node: StructureSource; children: readonly XmlContent[] }[] = [];
  for (let i = 0; i < count; i++) {
    const patch = patches[i];
    if (!patch || typeof patch !== 'object') fail('XML_STRUCTURE_INPUT', 'Expected a structured patch');
    const { target, children } = patch;
    const node = target && owned.get(target);
    if (!node) fail('XML_STRUCTURE_TARGET', 'Target is foreign, forged or absent');
    if (replace && !node.parent) fail('XML_STRUCTURE_ROOT', 'The root cannot be replaced');
    arrayLength(children);
    selected.push({ node, children });
  }
  selected.sort((a, b) => a.node.start - b.node.start);
  let previousEnd = -1;
  for (const { node } of selected) {
    if (node.start < previousEnd) fail('XML_STRUCTURE_OVERLAP', 'Targets must be unique and disjoint');
    previousEnd = node.end;
  }
  const retained = source.length - (replace ? selected.reduce((n, p) => n + p.node.end - p.node.start, 0) : 0);
  const renderer = new Renderer(MAX_SIZE - retained);
  const edits: { start: number; end: number; value: string }[] = [];
  for (const { node, children } of selected) {
    const parent = replace ? node.parent! : node;
    const content = renderer.render(children, scope(parent, renderer), parent.depth);
    if (replace) edits.push({ start: node.start, end: node.end, value: content });
    else if (content.length) {
      if (node.selfClosing) {
        const close = `></${node.name}>`;
        renderer.reserve(close.length - 2);
        edits.push({ start: node.openEnd - 2, end: node.openEnd, value: '>' + content + `</${node.name}>` });
      } else edits.push({ start: node.closeStart, end: node.closeStart, value: content });
    }
  }
  if (!edits.length) return source;
  const parts: string[] = []; let cursor = 0;
  for (const edit of edits) { parts.push(source.slice(cursor, edit.start), edit.value); cursor = edit.end; }
  parts.push(source.slice(cursor)); const output = parts.join('');
  try { validateXml(output); }
  catch (error) { fail('XML_STRUCTURE_UNSAFE', `Structured edit produced invalid XML: ${error instanceof Error ? error.message : String(error)}`); }
  return output;
}

function scope(node: StructureSource, renderer: Renderer): Map<string, string> {
  const ancestors: StructureSource[] = [];
  for (let current: StructureSource | undefined = node; current; current = current.parent) ancestors.push(current);
  const frame = new Map([['xml', XML]]);
  for (let i = ancestors.length - 1; i >= 0; i--) for (const [name, uri] of Object.entries(ancestors[i]!.attributes)) {
    renderer.scopeWork(1);
    if (name === 'xmlns') frame.set('', uri);
    else if (name.startsWith('xmlns:')) frame.set(name.slice(6), uri);
  }
  return frame;
}

class Renderer {
  private used = 0;
  private input = 0;
  private items = 0;
  private scopeVisits = 0;
  scopeWork(count: number) {
    this.scopeVisits += count;
    if (this.scopeVisits > 1_000_000) fail('XML_STRUCTURE_LIMIT', 'Namespace scope work exceeds limit');
  }
  private active = new Set<object>();
  constructor(private readonly capacity: number) {}
  reserve(size: number) {
    this.used += size;
    if (this.used > this.capacity) fail('XML_STRUCTURE_LIMIT', 'Rendered output exceeds XML size limit');
  }
  private field(value: unknown): string {
    if (typeof value !== 'string') fail('XML_STRUCTURE_INPUT', 'Expected a string field');
    this.input += value.length;
    if (this.input > MAX_SIZE) fail('XML_STRUCTURE_LIMIT', 'Structured input exceeds size limit');
    return value;
  }
  private escaped(value: string, attribute: boolean): string {
    let length = 0;
    for (const char of value) length += char === '&' ? 5 : char === '<' || char === '>' ? 4 : char === '\r' ? 5 : attribute && (char === '"' || char === "'") ? 6 : attribute && (char === '\t' || char === '\n') ? 5 : char.length;
    this.reserve(length);
    return attribute ? escapeAttribute(value) : escapeText(value);
  }
  private name(value: XmlExpandedName): XmlExpandedName {
    if (!value || typeof value !== 'object') fail('XML_STRUCTURE_NAME', 'Expected an expanded name');
    const localName = this.field(value.localName), namespaceURI = this.field(value.namespaceURI);
    let valid = false;
    if (!localName.includes(':')) try {
      inspectXmlEvents(`<${localName}/>`, { start(n) { valid = n.name === localName; }, end() {}, text() {}, comment() {}, instruction() {} });
    } catch { valid = false; }
    if (!valid || namespaceURI === XMLNS) fail('XML_STRUCTURE_NAME', 'Invalid local name or reserved namespace');
    // Reused URIs came from validated source bindings. New URIs pass through
    // budgeted escaping when their declarations are emitted.
    return { localName, namespaceURI };
  }
  render(children: readonly XmlContent[], inherited: Map<string, string>, depth: number): string {
    const length = arrayLength(children), parts: string[] = [];
    for (let i = 0; i < length; i++) this.content(children[i]!, inherited, depth, parts);
    return parts.join('');
  }
  private content(content: XmlContent, inherited: Map<string, string>, depth: number, parts: string[]): void {
    if (++this.items > MAX_ITEMS) fail('XML_STRUCTURE_LIMIT', 'Too many structured nodes or attributes');
    if (typeof content === 'string') { parts.push(this.escaped(this.field(content), false)); return; }
    if (!content || typeof content !== 'object') fail('XML_STRUCTURE_INPUT', 'Expected text or a structured element');
    if (this.active.has(content)) fail('XML_STRUCTURE_CYCLE', 'Cyclic structured content');
    if (depth >= MAX_DEPTH) fail('XML_STRUCTURE_LIMIT', 'Structured content exceeds XML depth limit');
    this.active.add(content);
    try {
      const { name, attributes = [], children = [] } = content;
      const element = this.name(name);
      this.scopeWork(inherited.size);
      const frame = new Map(inherited), declarations: string[] = [];
      const prefixes = new Map<string, string>();
      for (const [prefix, uri] of frame) if (prefix && !prefixes.has(uri)) prefixes.set(uri, prefix);
      let nextPrefix = 1;
      const lexical = (name: XmlExpandedName, attr: boolean): string => {
        const { localName, namespaceURI } = name;
        if (!namespaceURI) {
          if (attr && localName === 'xmlns') fail('XML_STRUCTURE_NAME', 'Namespace declarations are not authored attributes');
          if (!attr && (frame.get('') ?? '') !== '') { frame.set('', ''); declarations.push(' xmlns=""'); this.reserve(9); }
          return localName;
        }
        if (namespaceURI === XML) return 'xml:' + localName;
        if (!attr && frame.get('') === namespaceURI) return localName;
        const existing = prefixes.get(namespaceURI);
        if (existing) return existing + ':' + localName;
        while (frame.has('n' + nextPrefix)) nextPrefix++;
        const prefix = 'n' + nextPrefix++; frame.set(prefix, namespaceURI); prefixes.set(namespaceURI, prefix);
        const escaped = this.escaped(namespaceURI, true);
        const before = ` xmlns:${prefix}="`; this.reserve(before.length + 1);
        declarations.push(before + escaped + '"');
        return prefix + ':' + localName;
      };
      const qname = lexical(element, false), attrs: string[] = [], seen = new Set<string>();
      const count = arrayLength(attributes);
      for (let i = 0; i < count; i++) {
        if (++this.items > MAX_ITEMS) fail('XML_STRUCTURE_LIMIT', 'Too many structured nodes or attributes');
        const attr = attributes[i]; if (!attr || typeof attr !== 'object') fail('XML_STRUCTURE_INPUT', 'Expected a structured attribute');
        const name = this.name(attr.name), value = this.field(attr.value), key = name.namespaceURI + '\u0000' + name.localName;
        if (seen.has(key)) fail('XML_STRUCTURE_ATTRIBUTE', 'Duplicate expanded attribute'); seen.add(key);
        const q = lexical(name, true), escaped = this.escaped(value, true);
        this.reserve(q.length + 4); attrs.push(` ${q}="${escaped}"`);
      }
      const childCount = arrayLength(children);
      this.reserve(qname.length + (childCount ? qname.length + 5 : 3));
      parts.push('<', qname);
      for (const declaration of declarations) parts.push(declaration);
      for (const attr of attrs) parts.push(attr);
      if (!childCount) parts.push('/>');
      else {
        parts.push('>');
        for (let i = 0; i < childCount; i++) this.content(children[i]!, frame, depth + 1, parts);
        parts.push('</', qname, '>');
      }
    } finally { this.active.delete(content); }
  }
}
