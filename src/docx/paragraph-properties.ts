import { OoxmlError } from '../errors.ts';
import { applyEdits, attribute, escapeAttribute, type XmlElement } from '../xml/index.ts';
import { directParagraphStyle, replaceParagraphStyle, PARAGRAPH_PROPERTY_ORDER } from './paragraph-style.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', XMLNS = 'http://www.w3.org/2000/xmlns/';
export type ParagraphAlignment = 'left' | 'center' | 'right' | 'both';
export interface DirectParagraphProperties {
  alignment: ParagraphAlignment | null;
  spacingBefore: number | null;
  spacingAfter: number | null;
  keepLines: boolean | null;
  pageBreakBefore: boolean | null;
  widowControl: boolean | null;
}
export type ParagraphPropertiesPatch = Partial<DirectParagraphProperties>;
const keys = ['alignment', 'spacingBefore', 'spacingAfter', 'keepLines', 'pageBreakBefore', 'widowControl'] as const;
function fail(message: string): never { throw new OoxmlError('docx-paragraph-properties-unsupported', message); }
const word = (n: XmlElement, name: string) => n.namespaceURI === W && n.localName === name;
const alignments = ['left', 'center', 'right', 'both'];
function twips(value: unknown): number { if (typeof value !== 'number' || !Number.isInteger(value) || value < 0 || value > 31680) fail('Spacing must be integer twips from 0 through 31680'); return value; }
export function normalizeParagraphProperties(patch: ParagraphPropertiesPatch): ParagraphPropertiesPatch {
  if (!patch || typeof patch !== 'object' || ![Object.prototype, null].includes(Object.getPrototypeOf(patch))) fail('Expected a plain paragraph-property patch');
  const result: ParagraphPropertiesPatch = {};
  for (const key of Reflect.ownKeys(patch)) {
    if (typeof key !== 'string' || !keys.includes(key as typeof keys[number])) fail('Unknown paragraph-property field');
    const descriptor = Object.getOwnPropertyDescriptor(patch, key)!;
    if (!('value' in descriptor)) fail('Paragraph properties must be plain values');
    const value = descriptor.value;
    if (value === undefined) continue;
    if (value !== null) {
      if (key === 'alignment') { if (!alignments.includes(value)) fail('Unsupported alignment'); }
      else if (key.startsWith('spacing')) twips(value);
      else if (typeof value !== 'boolean') fail('Paragraph flags require boolean or null');
    }
    Object.assign(result, { [key]: value });
  }
  if (!Object.keys(result).length) fail('Expected at least one paragraph-property value');
  return result;
}
function inspect(xml: string, p: XmlElement) {
  directParagraphStyle(xml, p); // Unique, first, ordered pPr; no unknown/revised children.
  const pr = p.children.find(n => word(n, 'pPr'));
  if (pr) for (const name of Object.keys(pr.attributes)) if (pr.attributeNamespaces[name] !== XMLNS) fail('Paragraph-property attributes are unsupported');
  const properties: DirectParagraphProperties = { alignment: null, spacingBefore: null, spacingAfter: null, keepLines: null, pageBreakBefore: null, widowControl: null };
  const nodes = new Map<string, XmlElement>(), spacing = new Map<string, string>();
  for (const node of pr?.children ?? []) {
    if (!['jc', 'spacing', 'keepLines', 'pageBreakBefore', 'widowControl'].includes(node.localName)) continue;
    nodes.set(node.localName, node);
    if (node.children.length || (!node.selfClosing && !/^[ \t\r\n]*$/.test(xml.slice(node.openEnd, node.closeStart)))) fail('Selected paragraph properties must be empty leaves');
    const allowed = node.localName === 'spacing' ? ['before', 'after', 'line', 'lineRule', 'beforeAutospacing', 'afterAutospacing'] : ['val'];
    for (const name of Object.keys(node.attributes)) {
      if (node.attributeNamespaces[name] === XMLNS) continue;
      const local = name.split(':').at(-1)!;
      if (node.attributeNamespaces[name] !== W || !allowed.includes(local)) fail('Unsupported selected property attribute or namespace');
      if (node.localName === 'spacing') spacing.set(local, node.attributes[name]!);
    }
    const value = attribute(node, 'val', W);
    if (node.localName === 'jc') { if (!value || !alignments.includes(value)) fail('Unsupported direct alignment'); properties.alignment = value as ParagraphAlignment; }
    else if (node.localName !== 'spacing') {
      if (value !== undefined && !['true', 'false', 'on', 'off', '1', '0'].includes(value)) fail('Invalid direct on/off value');
      properties[node.localName as 'keepLines' | 'pageBreakBefore' | 'widowControl'] = value === undefined || ['true', 'on', '1'].includes(value);
    }
  }
  for (const name of ['before', 'after'] as const) {
    const value = spacing.get(name);
    if (value !== undefined) { if (!/^\d+$/.test(value)) fail('Invalid direct spacing'); properties[name === 'before' ? 'spacingBefore' : 'spacingAfter'] = twips(Number(value)); }
    const auto = spacing.get(name + 'Autospacing');
    if (auto !== undefined && !['false', 'off', '0'].includes(auto)) fail('Automatic spacing takes precedence over direct twips');
  }
  if (spacing.has('line') && !/^-?\d+$/.test(spacing.get('line')!)) fail('Invalid line spacing');
  if (spacing.has('lineRule') && !['auto', 'atLeast', 'exact'].includes(spacing.get('lineRule')!)) fail('Invalid line spacing rule');
  return { pr, nodes, spacing, properties };
}
export function readParagraphProperties(xml: string, p: XmlElement): DirectParagraphProperties { return inspect(xml, p).properties; }
export function editParagraphProperties(xml: string, p: XmlElement, patch: ParagraphPropertiesPatch): string {
  const { pr, nodes, spacing, properties } = inspect(xml, p);
  // Reuse existing plain-run/mixed-content/revision guards without changing style.
  replaceParagraphStyle(xml, p, directParagraphStyle(xml, p) ?? null);
  const edits: { start: number; end: number; value: string }[] = [], additions: { name: string; xml: string }[] = [];
  const changed = keys.filter(k => patch[k] !== undefined && patch[k] !== properties[k]);
  if (!changed.length) return xml;
  const leaf = (name: string, value: string | null) => value === null ? '' : `<w:${name} xmlns:w="${W}" w:val="${escapeAttribute(value)}"/>`;
  const replacements = new Map<string, string>();
  for (const key of changed) {
    const value = patch[key]!;
    if (key.startsWith('spacing')) {
      const name = key === 'spacingBefore' ? 'before' : 'after';
      if (value === null) spacing.delete(name); else spacing.set(name, String(value));
    } else replacements.set(key === 'alignment' ? 'jc' : key, leaf(key === 'alignment' ? 'jc' : key, value === null ? null : typeof value === 'boolean' ? value ? '1' : '0' : String(value)));
  }
  if (changed.some(k => k.startsWith('spacing'))) replacements.set('spacing', spacing.size ? `<w:spacing xmlns:w="${W}"${[...spacing].map(([name, value]) => ` w:${name}="${escapeAttribute(value)}"`).join('')}/>` : '');
  for (const [name, value] of replacements) {
    const node = nodes.get(name);
    if (node) edits.push({ start: node.start, end: node.end, value }); else if (value) additions.push({ name, xml: value });
  }
  additions.sort((a, b) => PARAGRAPH_PROPERTY_ORDER.indexOf(a.name) - PARAGRAPH_PROPERTY_ORDER.indexOf(b.name));
  if (!pr || pr.selfClosing) {
    const children = additions.map(a => a.xml).join('');
    if (pr) edits.push({ start: pr.start, end: pr.end, value: xml.slice(pr.start, pr.end).replace(/\/>$/, '>') + children + `</${pr.name}>` });
    else {
      const value = `<w:pPr xmlns:w="${W}">${children}</w:pPr>`;
      if (p.selfClosing) edits.push({ start: p.start, end: p.end, value: xml.slice(p.start, p.end).replace(/\/>$/, '>') + value + `</${p.name}>` });
      else edits.push({ start: p.openEnd, end: p.openEnd, value });
    }
  } else {
    const positions = new Map<number, string>();
    for (const addition of additions) {
      const next = pr.children.find(n => PARAGRAPH_PROPERTY_ORDER.indexOf(n.localName) > PARAGRAPH_PROPERTY_ORDER.indexOf(addition.name));
      const at = next?.start ?? pr.closeStart; positions.set(at, (positions.get(at) ?? '') + addition.xml);
    }
    for (const [start, value] of positions) edits.push({ start, end: start, value });
  }
  return applyEdits(xml, edits);
}
