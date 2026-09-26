import { OoxmlError } from '../errors.ts';
import { escapeAttribute, inspectXmlEvents, validateXml } from './index.ts';
import type { XmlRemovalTarget } from './removal.ts';

export interface XmlAttributePatch { readonly target: XmlRemovalTarget; readonly name: string; readonly value: string }
export interface AttributeSource {
  start: number; openEnd: number; selfClosing: boolean;
  attributes: Record<string, string>; namespaces: Readonly<Record<string, string>>;
  parent?: AttributeSource;
}
type Span = { name: string; start: number; end: number; quote: string };
const MAX_OUTPUT = 8 * 1024 * 1024, MAX_PATCHES = 100_000;
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
function fail(code: string, message: string): never { throw new OoxmlError(code, message); }

export function setSnapshotAttributes(source: string, patches: readonly XmlAttributePatch[], owned: WeakMap<XmlRemovalTarget, AttributeSource>): string {
  if (!Array.isArray(patches)) fail('XML_ATTRIBUTE_PATCH', 'Expected an array of attribute patches');
  const count = patches.length;
  if (!Number.isSafeInteger(count) || count < 0 || count > MAX_PATCHES) fail('XML_ATTRIBUTE_LIMIT', 'Attribute patch count exceeds limit');
  const groups = new Map<AttributeSource, { seen: Set<string>; spans?: Map<string, Span>; additions: string[] }>();
  const edits: { start: number; end: number; value: string }[] = [];
  let outputLength = source.length, inputLength = 0;
  for (let i = 0; i < count; i++) {
    const patch = patches[i];
    if (!patch || typeof patch !== 'object') fail('XML_ATTRIBUTE_PATCH', 'Expected an attribute patch object');
    const { target, name, value } = patch;
    const node = target && owned.get(target);
    if (!node) fail('XML_ATTRIBUTE_TARGET', 'Attribute target is foreign, forged or absent');
    if (typeof name !== 'string' || typeof value !== 'string') fail('XML_ATTRIBUTE_PATCH', 'Attribute name and value must be strings');
    inputLength += name.length + value.length;
    if (inputLength > MAX_OUTPUT) fail('XML_ATTRIBUTE_LIMIT', 'Attribute patch input exceeds limit');
    const [local, namespace] = expandedName(name, node);
    const key = namespace + '\u0000' + local;
    let group = groups.get(node);
    if (!group) { group = { seen: new Set(), additions: [] }; groups.set(node, group); }
    if (group.seen.has(key)) fail('XML_ATTRIBUTE_DUPLICATE', 'Duplicate expanded attribute in one batch');
    group.seen.add(key);
    group.spans ??= new Map(attributeSpans(source, node).map(span => [node.namespaces[span.name] + '\u0000' + span.name.split(':').at(-1), span]));
    const existing = group.spans.get(key);
    if (existing && node.attributes[existing.name] === value) continue;
    const quote = existing?.quote ?? '"';
    // Count exact escaped growth before allocating replacement strings.
    let size = 0;
    for (const char of value) {
      size += char === '&' ? 5 : char === '<' || char === '>' ? 4 : char === quote ? (quote === '"' ? 6 : 5) : /[\t\r\n]/.test(char) ? 5 : char.length;
    }
    outputLength += size - (existing ? existing.end - existing.start : 0) + (existing ? 0 : name.length + 4);
    if (outputLength > MAX_OUTPUT) fail('XML_ATTRIBUTE_LIMIT', 'Edited XML exceeds input-size limit');
    const escaped = quote === "'" ? escapeAttribute(value).replaceAll('&apos;', '&#39;').replaceAll('&quot;', '"') : escapeAttribute(value).replaceAll('&apos;', "'");
    if (existing) edits.push({ start: existing.start, end: existing.end, value: escaped });
    else group.additions.push(` ${name}="${escaped}"`);
  }
  for (const [node, group] of groups) if (group.additions.length) {
    const at = node.openEnd - (node.selfClosing ? 2 : 1);
    edits.push({ start: at, end: at, value: group.additions.join('') });
  }
  if (!edits.length) return source;
  edits.sort((a, b) => a.start - b.start);
  const parts: string[] = []; let cursor = 0;
  for (const edit of edits) { parts.push(source.slice(cursor, edit.start), edit.value); cursor = edit.end; }
  parts.push(source.slice(cursor)); const output = parts.join('');
  try { validateXml(output); }
  catch (error) { fail('XML_ATTRIBUTE_UNSAFE', `Attribute edits produced invalid XML: ${error instanceof Error ? error.message : String(error)}`); }
  return output;
}

function expandedName(name: string, node: AttributeSource): [string, string] {
  const pieces = name.split(':');
  if (pieces.length > 2 || name === 'xmlns' || pieces[0] === 'xmlns') fail('XML_ATTRIBUTE_NAME', 'Namespace declarations cannot be edited');
  for (const piece of pieces) {
    let valid = false;
    try { inspectXmlEvents(`<${piece}/>`, { start(n) { valid = n.name === piece && !piece.includes(':'); }, end() {}, text() {}, comment() {}, instruction() {} }); }
    catch { valid = false; }
    if (!valid) fail('XML_ATTRIBUTE_NAME', 'Invalid attribute QName');
  }
  if (pieces.length === 1) return [name, ''];
  const [prefix, local] = pieces;
  if (prefix === 'xml') return [local!, XML_NS];
  for (let current: AttributeSource | undefined = node; current; current = current.parent) {
    if (Object.hasOwn(current.attributes, 'xmlns:' + prefix)) return [local!, current.attributes['xmlns:' + prefix]!];
  }
  return fail('XML_ATTRIBUTE_NAME', 'Unbound attribute prefix');
}

function attributeSpans(source: string, node: AttributeSource): Span[] {
  // The source has already passed the XML scanner. This pass only locates values.
  const whitespace = (c: string) => c === ' ' || c === '\t' || c === '\r' || c === '\n';
  let cursor = node.start + 1;
  while (cursor < node.openEnd && !whitespace(source[cursor]!) && source[cursor] !== '/' && source[cursor] !== '>') cursor++;
  const spans: Span[] = [];
  while (cursor < node.openEnd) {
    while (whitespace(source[cursor]!)) cursor++;
    if (source[cursor] === '/' || source[cursor] === '>') break;
    const nameStart = cursor;
    while (!whitespace(source[cursor]!) && source[cursor] !== '=') cursor++;
    const name = source.slice(nameStart, cursor);
    while (whitespace(source[cursor]!)) cursor++;
    cursor++; // equals
    while (whitespace(source[cursor]!)) cursor++;
    const quote = source[cursor++]!, start = cursor;
    while (source[cursor] !== quote) cursor++;
    spans.push({ name, quote, start, end: cursor });
    cursor++;
  }
  return spans;
}
