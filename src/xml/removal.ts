import { OoxmlError } from '../errors.ts';
import { inspectXmlEvents, validateXml, type XmlElement } from './index.ts';
import { setSnapshotAttributes, type AttributeSource, type XmlAttributePatch } from './attributes.ts';
export type { XmlAttributePatch } from './attributes.ts';

/** Immutable UTF-16 source offsets. Only the issuing snapshot accepts this handle. */
export interface XmlRemovalTarget {
  readonly name: string;
  readonly localName: string;
  readonly namespaceURI: string;
  readonly start: number;
  readonly end: number;
}

/** A reusable, syntax-checked string snapshot for disjoint subtree removal. */
export class XmlSnapshot {
  readonly elements: readonly XmlRemovalTarget[];
  readonly #source: string;
  readonly #owned = new WeakSet<XmlRemovalTarget>();
  readonly #attributeSources = new WeakMap<XmlRemovalTarget, AttributeSource>();

  private constructor(source: string) {
    if (typeof source !== 'string') fail('XML_REMOVAL_SOURCE', 'Expected an XML source string');
    this.#source = source;
    // The scanner supplies end offsets after closing each element. Retain nodes
    // only until scanning ends; never request accumulated descendant text.
    const nodes: XmlElement[] = [];
    inspectXmlEvents(source, {
      start(node) { nodes.push(node); }, end() {}, text() {}, comment() {}, instruction() {},
    });
    const metadata = new Map<XmlElement, AttributeSource>();
    this.elements = Object.freeze(nodes.map(node => {
      const { name, localName, namespaceURI, start, end } = node;
      const target = Object.freeze({ name, localName, namespaceURI, start, end });
      const attributes: AttributeSource = { start, openEnd: node.openEnd, selfClosing: node.selfClosing,
        attributes: node.attributes, namespaces: node.attributeNamespaces, parent: node.parent ? metadata.get(node.parent) : undefined };
      metadata.set(node, attributes);
      this.#owned.add(target);
      this.#attributeSources.set(target, attributes);
      return target;
    }));
    Object.freeze(this);
  }

  static parse(source: string): XmlSnapshot { return new XmlSnapshot(source); }

  /** Set existing or new attributes in the original snapshot's namespace scope. */
  setAttributes(patches: readonly XmlAttributePatch[]): string {
    return setSnapshotAttributes(this.#source, patches, this.#attributeSources);
  }

  /** Return a new string; the original snapshot and all of its handles stay valid. */
  remove(targets: readonly XmlRemovalTarget[]): string {
    if (!Array.isArray(targets)) fail('XML_REMOVAL_TARGET', 'Expected an array of snapshot-owned targets');
    const count = targets.length;
    if (!Number.isSafeInteger(count) || count < 0 || count > this.elements.length) {
      fail('XML_REMOVAL_TARGET', 'Expected a bounded array of snapshot-owned targets');
    }
    const selected: XmlRemovalTarget[] = [];
    // Array iterators are caller-controlled; read only the bounded indexed entries.
    for (let i = 0; i < count; i++) {
      const target = targets[i];
      if (!target || !this.#owned.has(target)) fail('XML_REMOVAL_TARGET', 'Target belongs to another snapshot or is forged');
      if (target === this.elements[0]) fail('XML_REMOVAL_ROOT', 'The document root cannot be removed');
      selected.push(target);
    }
    selected.sort((a, b) => a.start - b.start);
    let previousEnd = -1;
    for (const target of selected) {
      if (target.start < previousEnd) fail('XML_REMOVAL_OVERLAP', 'Removal targets must be unique and disjoint');
      previousEnd = target.end;
    }
    if (!selected.length) return this.#source;
    const parts: string[] = [];
    let cursor = 0;
    for (const target of selected) {
      parts.push(this.#source.slice(cursor, target.start));
      cursor = target.end;
    }
    parts.push(this.#source.slice(cursor));
    const output = parts.join('');
    // Adjacent text can become illegal even when removing whole elements:
    // removing <a/> from ]]<a/>> creates the forbidden text sequence ]]>.
    try { validateXml(output); }
    catch (error) {
      throw new OoxmlError('XML_REMOVAL_UNSAFE', `Removal produced invalid XML: ${error instanceof Error ? error.message : String(error)}`);
    }
    return output;
  }
}

function fail(code: string, message: string): never { throw new OoxmlError(code, message); }
