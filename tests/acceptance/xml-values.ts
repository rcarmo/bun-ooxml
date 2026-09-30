import assert from 'node:assert/strict';
import type { AcceptanceStep, StepBinding } from '../../scripts/gherkin.ts';
import { OoxmlError, classifyXmlParseFailure } from '../../src/errors.ts';
import { attribute, escapeAttribute, escapeText, parseXml, type XmlDocument } from '../../src/xml/index.ts';

type State = { source?: string; originalSource?: string; value?: string; document?: XmlDocument; escaped?: string; error?: unknown };
const state = (c: Record<string, unknown>) => c.state as State;
function root(c: Record<string, unknown>) { const s = (c.lexical as State | undefined) ?? state(c); assert.equal(s.error, undefined); assert(s.document); return s.document.root; }
function jsonString(source: string | undefined) { assert(source !== undefined); const value: unknown = JSON.parse(source); assert.equal(typeof value, 'string'); return value as string; }
export const scenarioIds = [
  '@id-xml-entity-values', '@id-xml-stylesheet-processing-instruction',
  '@id-xml-expanded-attribute-lookup', '@id-xml-implicit-xml-prefix',
  '@id-xml-prototype-safe-attributes', '@id-xml-immutable-namespace-metadata',
  '@id-xml-escaping-values', '@id-xml-escaping-invalid-character',
  '@id-xml-escaping-whitespace-roundtrip', '@id-xml-typed-parse-error',
];
export const bindings: StepBinding[] = [
  { pattern: /^XML values input encoded as JSON (.+)$/, run: (c, json) => { state(c).source = jsonString(json); state(c).originalSource = state(c).source; } },
  { pattern: /^the XML values input is parsed$/, run: c => {
    const s = state(c); assert(s.source !== undefined);
    try { s.document = parseXml(s.source); } catch (error) { s.error = error; }
  } },
  { pattern: /^the root attribute a equals JSON (.+) and the child expanded attribute urn:x\/b equals JSON (.+)$/, run: (c,a,b) => {const r=root(c);assert.equal(r.attributes.a,jsonString(a));assert.equal(attribute(r.children[0]!,'b','urn:x'),jsonString(b));} },
  { pattern: /^the root attribute (\S+) equals JSON ((?!.* and the child expanded attribute).+)$/, run: (c, name, json) => {
    const r = root(c); assert(Object.hasOwn(r.attributes, name!)); assert.equal(r.attributes[name!], jsonString(json));
  } },
  { pattern: /^the root text equals JSON (.+)$/, run: (c, json) => { assert.equal(root(c).text, jsonString(json)); } },
  { pattern: /^the root qualified name equals (\S+)$/, run: (c, name) => { assert.equal(root(c).name, name); } },
  { pattern: /^the root namespace URI equals JSON (.+)$/, run: (c, json) => { assert.equal(root(c).namespaceURI, jsonString(json)); } },
  { pattern: /^the implicit xml namespace URI is (\S+)$/, run: (c, uri) => {
    const r = root(c); assert.equal(r.attributeNamespaces['xml:lang'], uri); assert.equal(attribute(r, 'lang', uri), 'en');
  } },
  { pattern: /^expanded attribute lookups return these JSON values$/, run: c => {
    const r = root(c), rows = (c.step as AcceptanceStep).argument?.dataTable;
    assert(rows && rows.length > 1); assert.deepEqual(rows[0], ['element', 'local', 'namespace', 'value_json']);
    for (const row of rows.slice(1)) {
      assert.equal(row.length, 4);
      const [element, local, namespace, value] = row;
      const nodes = element === 'root' ? [r] : r.children.filter(n => n.localName === element);
      assert.equal(nodes.length, 1);
      const expected: unknown = JSON.parse(value!); assert(expected === null || typeof expected === 'string');
      assert.equal(attribute(nodes[0]!, local!, namespace!), expected === null ? undefined : expected);
    }
  } },
  { pattern: /^the root attributes are exactly __proto__=polluted and constructor=safe$/, run: c => {
    const attrs = root(c).attributes;
    assert.deepEqual(Object.keys(attrs).sort(), ['__proto__', 'constructor']);
    assert(Object.hasOwn(attrs, '__proto__')); assert(Object.hasOwn(attrs, 'constructor'));
    assert.equal(attrs.__proto__, 'polluted'); assert.equal(attrs.constructor, 'safe');
    // Retain Bun's native guard in addition to the portable outcomes.
    assert.equal(Object.getPrototypeOf(attrs), null);
  } },
  { pattern: /^the root remains r with no text and no children$/, run: c => {
    const r = root(c); assert.equal(r.name, 'r'); assert.equal(r.text, ''); assert.deepEqual(r.children, []);
  } },
  { pattern: /^fresh attribute reads and the original XML source are unchanged$/, run: c => {
    const r = root(c); assert.equal(attribute(r, '__proto__'), 'polluted'); assert.equal(attribute(r, 'constructor'), 'safe');
    assert.equal(state(c).source, state(c).originalSource);
  } },
  { pattern: /^changing a returned namespace snapshot to urn:changed and deleting a:id are attempted$/, run: c => {
    const r = root(c), snapshot = r.attributeNamespaces;
    const check = () => { assert.equal(root(c).attributeNamespaces['a:id'], 'urn:a'); assert.equal(attribute(root(c), 'id', 'urn:a'), 'outer'); };
    // Each attempt may refuse or affect a detached copy, never the original model.
    Reflect.set(snapshot, 'a:id', 'urn:changed'); check();
    Reflect.deleteProperty(snapshot, 'a:id'); check();
    assert.equal(r.name, 'r'); assert.equal(r.text, ''); assert.deepEqual(r.children, []);
    assert.equal(state(c).source, state(c).originalSource);
  } },
  { pattern: /^fresh namespace and attribute reads still return urn:a and outer$/, run: c => {
    assert.equal(root(c).attributeNamespaces['a:id'], 'urn:a'); assert.equal(attribute(root(c), 'id', 'urn:a'), 'outer');
  } },
  { pattern: /^the parsed root structure and original XML source are unchanged$/, run: c => {
    const r = root(c); assert.equal(r.name, 'r'); assert.equal(r.text, ''); assert.deepEqual(r.children, []);
    assert.equal(state(c).source, state(c).originalSource);
  } },
  { pattern: /^parsing fails with category malformed-xml and no document result$/, run: c => {
    const s = state(c); assert.equal(classifyXmlParseFailure(s.error), 'malformed-xml'); assert.equal(s.document, undefined);
  } },
  { pattern: /^the original XML source is unchanged$/, run: c => { assert.equal(state(c).source, state(c).originalSource); } },
  { pattern: /^the root attribute map has a null prototype$/, run: c => { assert.equal(Object.getPrototypeOf(root(c).attributes), null); } },
  { pattern: /^(\S+) is an own attribute with value (\S+)$/, run: (c, name, value) => {
    const attrs = root(c).attributes; assert(Object.hasOwn(attrs, name!)); assert.equal(attrs[name!], value);
  } },
  { pattern: /^the namespace recorded for (\S+) is (\S+)$/, run: (c, name, uri) => { assert.equal(root(c).attributeNamespaces[name!], uri); } },
  { pattern: /^the attribute namespace map is frozen and has a null prototype$/, run: c => {
    const map = root(c).attributeNamespaces; assert(Object.isFrozen(map)); assert.equal(Object.getPrototypeOf(map), null);
  } },
  { pattern: /^an XML escaping value encoded as JSON (.+)$/, run: (c, json) => { state(c).value = jsonString(json); } },
  { pattern: /^the value is escaped for XML (text|attribute) content$/, run: (c, context) => {
    const s = state(c); assert(s.value !== undefined);
    try { s.escaped = context === 'text' ? escapeText(s.value) : escapeAttribute(s.value); } catch (error) { s.error = error; }
  } },
  { pattern: /^the escaped string equals JSON (.+)$/, run: (c, json) => {
    const s = state(c); assert.equal(s.error, undefined); assert.equal(s.escaped, jsonString(json));
  } },
  { pattern: /^escaping refuses the invalid XML character$/, run: c => {
    const s = state(c); assert(s.error instanceof OoxmlError); assert.equal(s.error.code, 'XML_INVALID_CHAR'); assert.equal(s.escaped, undefined);
  } },
  { pattern: /^the value is escaped separately as text and as an attribute and both are parsed$/, run: c => {
    const s = state(c); assert(s.value !== undefined);
    s.document = parseXml(`<r a="${escapeAttribute(s.value)}">${escapeText(s.value)}</r>`);
  } },
  { pattern: /^the decoded text and attribute both equal JSON (.+)$/, run: (c, json) => {
    const r = root(c), value = jsonString(json); assert.equal(r.text, value); assert.equal(r.attributes.a, value);
  } },
  { pattern: /^parsing throws an OoxmlError instance$/, run: c => {
    const s = state(c); assert(s.error instanceof OoxmlError); assert.equal(s.document, undefined);
  } },
];
