import { expect, test } from 'bun:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Document } from '../../src/index.ts';
import { OpcPackage, addPart, addRelationship } from '../../src/opc/index.ts';
import { parseXml, elements, attribute } from '../../src/xml/index.ts';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function fixture(properties = '', text = 'Alpha') {
  const doc = Document.create(); doc.addParagraph(text);
  const pkg = await OpcPackage.open(doc.package.toBytes()), main = pkg.mainPart();
  const xml = pkg.text(main), match = /<w:p(?:\s[^>]*)?>/.exec(xml); expect(match).not.toBeNull();
  pkg.set(main, xml.slice(0, match!.index + match![0].length) + properties + xml.slice(match!.index + match![0].length));
  return Document.open(pkg.toBytes());
}

test('direct paragraph properties set and read all supported values without changing text', async () => {
  const doc = await fixture(), before = doc.package.parts;
  expect(doc.paragraphs[0]!.directProperties()).toEqual({ alignment: null, spacingBefore: null, spacingAfter: null, keepLines: null, pageBreakBefore: null, widowControl: null, outlineLevel: null });
  expect(doc.paragraphs[0]!.setProperties({ alignment: 'both', spacingBefore: 240, spacingAfter: 120, keepLines: true, pageBreakBefore: false, widowControl: true })).toEqual({ changed: 1 });
  expect(doc.paragraphs[0]!.directProperties()).toEqual({ alignment: 'both', spacingBefore: 240, spacingAfter: 120, keepLines: true, pageBreakBefore: false, widowControl: true, outlineLevel: null });
  expect(doc.paragraphs[0]!.text).toBe('Alpha');
  const reopened = await Document.open(doc.package.toBytes());
  expect(reopened.paragraphs[0]!.directProperties()).toEqual(doc.paragraphs[0]!.directProperties());
  for (const [name, bytes] of before) if (name !== 'word/document.xml') expect(reopened.package.get(name)).toEqual(bytes);
  const xml = new TextDecoder().decode(doc.package.get('word/document.xml'));
  expect(elements(parseXml(xml), 'pPr', W)[0]!.children.map(n => n.localName)).toEqual(['keepLines', 'pageBreakBefore', 'widowControl', 'spacing', 'jc']);
});

test('direct values distinguish absent, explicit zero/off and removal; no-op preserves archive and handles', async () => {
  const doc = await fixture('<w:pPr><w:keepLines/><w:spacing w:before="0"/><w:jc w:val="left"/></w:pPr>'), p = doc.paragraphs[0]!, before = doc.package.toBytes();
  expect(p.directProperties()).toEqual({ alignment: 'left', spacingBefore: 0, spacingAfter: null, keepLines: true, pageBreakBefore: null, widowControl: null, outlineLevel: null });
  expect(p.setProperties({ keepLines: true, spacingBefore: 0, alignment: 'left' })).toEqual({ changed: 0 });
  expect(p.text).toBe('Alpha'); expect(doc.package.toBytes()).toEqual(before);
  expect(p.setProperties({ keepLines: null, spacingBefore: null, alignment: null })).toEqual({ changed: 1 });
  expect(doc.paragraphs[0]!.directProperties().keepLines).toBeNull();
  expect(doc.paragraphs[0]!.directProperties().spacingBefore).toBeNull();
  expect(doc.paragraphs[0]!.directProperties().alignment).toBeNull();
  expect(() => p.directProperties()).toThrow(expect.objectContaining({ code: 'docx-stale-paragraph' }));
});

test('spacing edits preserve unrelated line/auto attributes and other paragraph property source', async () => {
  const pr = '<w:pPr><w:keepNext/><w:spacing w:before = \'120\' w:after="240" w:line="360" w:lineRule="auto"/><w:ind w:left="720"/><w:jc w:val="right"/><w:rPr><w:b/></w:rPr></w:pPr>';
  const doc = await fixture(pr); doc.paragraphs[0]!.setProperties({ spacingBefore: 480, alignment: null });
  const xml = new TextDecoder().decode(doc.package.get('word/document.xml'));
  expect(xml).toContain('<w:keepNext/>'); expect(xml).toContain('<w:ind w:left="720"/>'); expect(xml).toContain('<w:rPr><w:b/></w:rPr>');
  const spacing = elements(parseXml(xml), 'spacing', W)[0]!;
  expect(attribute(spacing, 'line', W)).toBe('360'); expect(attribute(spacing, 'lineRule', W)).toBe('auto'); expect(attribute(spacing, 'after', W)).toBe('240');
  expect(doc.paragraphs[0]!.directProperties().spacingBefore).toBe(480);
});

test('namespace aliases and default Word namespaces retain expanded property names', async () => {
  const doc = await fixture('<w:pPr><w:spacing w:before="10"/></w:pPr>'), pkg = await OpcPackage.open(doc.package.toBytes());
  pkg.set(pkg.mainPart(), pkg.text(pkg.mainPart()).replaceAll('xmlns:w=', 'xmlns:q=').replaceAll('w:', 'q:'));
  const alias = await Document.open(pkg.toBytes()); alias.paragraphs[0]!.setProperties({ alignment: 'center', spacingAfter: 20 });
  expect(alias.paragraphs[0]!.directProperties().alignment).toBe('center'); expect(alias.paragraphs[0]!.directProperties().spacingBefore).toBe(10);
  const xml = parseXml(new TextDecoder().decode(alias.package.get('word/document.xml')));
  expect(elements(xml, 'jc', W)).toHaveLength(1); expect(attribute(elements(xml, 'spacing', W)[0]!, 'after', W)).toBe('20');
});

test('invalid or executable patch fields refuse atomically', async () => {
  const doc = await fixture(), before = doc.package.toBytes(), p = doc.paragraphs[0]!;
  const bad = [null, {}, { alignment: 'justify' }, { spacingBefore: -1 }, { spacingAfter: 1.5 }, { spacingAfter: Number.MAX_SAFE_INTEGER + 1 }, { keepLines: 1 }, { pageBreakBefore: 'true' }, { extra: true }, { alignment: 'left', keepLines: 'bad' }];
  let called = false; bad.push(Object.defineProperty({}, 'alignment', { enumerable: true, get() { called = true; return 'left'; } }));
  for (const patch of bad) { expect(() => p.setProperties(patch as never)).toThrow(); expect(doc.package.toBytes()).toEqual(before); }
  expect(called).toBe(false);
});

test('unknown, duplicated, revised, misordered or ambiguous selected properties refuse even on no-op', async () => {
  for (const pr of [
    '<w:pPr><w:jc w:val="left"/><w:jc w:val="left"/></w:pPr>',
    '<w:pPr><w:jc w:val="left"/><w:keepLines/></w:pPr>',
    '<w:pPr><w:pPrChange/></w:pPr>', '<w:pPr><!--barrier--><w:jc w:val="left"/></w:pPr>',
    '<w:pPr><w:jc val="left"/></w:pPr>', '<w:pPr><w:jc w:val="start"/></w:pPr>',
    '<w:pPr><w:spacing w:before="120" w:beforeLines="100"/></w:pPr>',
    '<w:pPr><w:spacing w:before="120" w:beforeAutospacing="1"/></w:pPr>',
    '<w:pPr><w:keepLines w:val="maybe"/></w:pPr>', '<w:pPr><w:spacing w:before="-1"/></w:pPr>',
    '<w:pPr><w:framePr/><w:mystery/></w:pPr>',
  ]) {
    const doc = await fixture(pr), before = doc.package.toBytes();
    expect(() => doc.paragraphs[0]!.setProperties({ alignment: 'left' })).toThrow(); expect(doc.package.toBytes()).toEqual(before);
  }
});

test('protection, unsupported text and externally changed document bytes refuse without mutation', async () => {
  const doc = await fixture(), pkg = await OpcPackage.open(doc.package.toBytes());
  addPart(pkg, 'word/settings.xml', `<w:settings xmlns:w="${W}"><w:documentProtection w:enforcement="1"/></w:settings>`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  addRelationship(pkg, pkg.mainPart(), 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings', 'settings.xml');
  const protectedDoc = await Document.open(pkg.toBytes()), before = protectedDoc.package.toBytes();
  expect(() => protectedDoc.paragraphs[0]!.setProperties({ keepLines: true })).toThrow(); expect(protectedDoc.package.toBytes()).toEqual(before);
  const held = doc.paragraphs[0]!; doc.package.setPart('word/document.xml', new TextEncoder().encode(new TextDecoder().decode(doc.package.get('word/document.xml')).replace('Alpha', 'Beta')));
  const changed = doc.package.toBytes(); expect(() => held.setProperties({ keepLines: true })).toThrow(); expect(doc.package.toBytes()).toEqual(changed);
  const other = await fixture(); other.package.setPart('word/document.xml', new TextEncoder().encode(new TextDecoder().decode(other.package.get('word/document.xml')).replace('<w:r>', '<w:hyperlink><w:r>').replace('</w:r>', '</w:r></w:hyperlink>')));
  const unsupported = await Document.open(other.package.toBytes()); expect(() => unsupported.paragraphs[0]!.setProperties({ keepLines: true })).toThrow();
});

test('disk save and UTF-16LE reopen retain properties and untouched parts', async () => {
  const doc = await fixture(), pkg = await OpcPackage.open(doc.package.toBytes());
  const text = pkg.text(pkg.mainPart()).replace('UTF-8', 'UTF-16'); const bytes = new Uint8Array(2 + text.length * 2); bytes[0] = 255; bytes[1] = 254;
  const view = new DataView(bytes.buffer); for (let i = 0; i < text.length; i++) view.setUint16(2 + 2 * i, text.charCodeAt(i), true);
  pkg.set(pkg.mainPart(), bytes); const utf = await Document.open(pkg.toBytes()), before = utf.package.parts;
  utf.paragraphs[0]!.setProperties({ spacingBefore: 240, keepLines: true });
  const root = await mkdtemp(join(tmpdir(), 'docx-ppr-'));
  try { const path = join(root, 'edited.docx'); await utf.save(path); const reopened = await Document.open(path);
    expect(reopened.paragraphs[0]!.directProperties().spacingBefore).toBe(240); expect(reopened.paragraphs[0]!.directProperties().keepLines).toBe(true);
    expect([...reopened.package.get('word/document.xml')!.slice(0, 2)]).toEqual([255, 254]);
    for (const [n, b] of before) if (n !== 'word/document.xml') expect(reopened.package.get(n)).toEqual(b);
  } finally { await rm(root, { recursive: true, force: true }); }
});

test('self-closing properties and empty paragraphs accept ordered insertion under a default namespace', async () => {
  for (const body of [`<p xmlns="${W}"/>`, `<p xmlns="${W}"><pPr/><r><t>Alpha</t></r></p>`, `<p xmlns="${W}"><pPr><spacing xmlns:q="${W}" q:line="360" q:afterAutospacing="0"/></pPr><r><t>Alpha</t></r></p>`]) {
    const d = Document.create(), pkg = await OpcPackage.open(d.package.toBytes());
    const xml = pkg.text(pkg.mainPart()).replace('<w:body>', '<w:body>' + body); pkg.set(pkg.mainPart(), xml);
    const doc = await Document.open(pkg.toBytes()); expect(doc.paragraphs).toHaveLength(1);
    doc.paragraphs[0]!.setProperties({ spacingAfter: 0, keepLines: false, alignment: 'right' });
    const values = doc.paragraphs[0]!.directProperties(); expect(values.spacingAfter).toBe(0); expect(values.keepLines).toBe(false); expect(values.alignment).toBe('right');
    expect((await Document.open(doc.package.toBytes())).paragraphs[0]!.directProperties()).toEqual(values);
  }
});

test('late serialization and staged write failures restore package, values and held paragraph', async () => {
  for (const stage of ['set', 'toBytes'] as const) {
    const doc = await fixture(), paragraph = doc.paragraphs[0]!, before = doc.package.toBytes();
    const pkg = (doc as unknown as { opcPackage: OpcPackage }).opcPackage;
    const original = pkg[stage].bind(pkg);
    if (stage === 'set') pkg.set = (name, value) => { (original as OpcPackage['set'])(name, value); throw new Error('injected write failure'); };
    else pkg.toBytes = () => { throw new Error('injected serialization failure'); };
    try { expect(() => paragraph.setProperties({ alignment: 'center' })).toThrow('injected'); }
    finally { if (stage === 'set') pkg.set = original as OpcPackage['set']; else pkg.toBytes = original as OpcPackage['toBytes']; }
    expect(doc.package.toBytes()).toEqual(before); expect(paragraph.text).toBe('Alpha'); expect(paragraph.directProperties().alignment).toBeNull();
    expect(paragraph.setProperties({ alignment: 'center' })).toEqual({ changed: 1 });
  }
});

test('table paragraph formatting preserves grid and live table access while expiring paragraph handles', async () => {
  const doc = Document.create(); doc.addTable(1, 2); doc.tables[0]!.cell(0, 0).text = 'cell'; doc.tables[0]!.cell(0, 1).text = 'other';
  const table = doc.tables[0]!, paragraph = doc.paragraphs[0]!;
  paragraph.setProperties({ spacingBefore: 120, keepLines: true });
  expect(table.rows).toBe(1); expect(table.columns).toBe(2); expect(table.cell(0, 0).text).toBe('cell');
  expect(() => paragraph.text).toThrow(expect.objectContaining({ code: 'docx-stale-paragraph' }));
  expect(doc.paragraphs[0]!.directProperties().spacingBefore).toBe(120);
  expect((await Document.open(doc.package.toBytes())).tables[0]!.cell(0, 1).text).toBe('other');
});

test('same-value requests still enforce protection and direct getters remain read-only', async () => {
  const doc = await fixture('<w:pPr><w:keepLines w:val="true"/></w:pPr>');
  const before = doc.package.toBytes(); expect(doc.paragraphs[0]!.directProperties().keepLines).toBe(true); expect(doc.package.toBytes()).toEqual(before);
  const pkg = await OpcPackage.open(before);
  addPart(pkg,'word/settings.xml',`<w:settings xmlns:w="${W}"><w:documentProtection/></w:settings>`,'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml');
  addRelationship(pkg,pkg.mainPart(),'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings','settings.xml');
  const p = await Document.open(pkg.toBytes()), protectedBytes = p.package.toBytes();
  expect(() => p.paragraphs[0]!.setProperties({ keepLines: true })).toThrow(expect.objectContaining({ code: 'docx-format-protected' }));
  expect(p.package.toBytes()).toEqual(protectedBytes);
});

test('nine shared direct paragraph getter cases execute with no credit for other document-model operations', async () => {
  const { fixturesRoot } = await import('../../scripts/fixture-inputs.ts'), { join } = await import('node:path');
  const { selectSharedScenarios, executeAcceptance } = await import('../../scripts/gherkin.ts');
  const { bindings } = await import('../acceptance/steps.ts'), { scenarioIds } = await import('../acceptance/paragraph-properties.ts');
  const path = 'workflows/docx/document-model.feature', text = await Bun.file(join(fixturesRoot(), path)).text();
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  const run = (source: string) => executeAcceptance({ root: '.', features: [selectSharedScenarios(path, source, scenarioIds)], counts: { features: count(1), scenarios: count(3), cases: count(9), steps: count(27) } }, bindings, 'paragraph-properties-unit');
  const good = await run(text); expect(good.failures).toEqual([]); expect(good.counts.cases.passed).toBe(9); expect(good.counts.cases.planned).toBeGreaterThan(0);
  for (const [from, to] of [['its alignment getter equals <value>', 'its alignment getter equals wrong'], ['its before and after getters equal <before> and <after>', 'its before and after getters equal 999 and 999']]) {
    expect(text.includes(from!)).toBe(true); const bad = await run(text.replace(from!, to!));
    expect(bad.counts.steps.failed).toBe(4); expect(bad.counts.steps.undefined).toBe(0); expect(bad.counts.steps.ambiguous).toBe(0);
  }
});

test('spacing policy boundaries and partial removal preserve an unedited line-spacing leaf', async () => {
  const doc = await fixture('<w:pPr><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/></w:pPr>');
  doc.paragraphs[0]!.setProperties({ spacingBefore: 31680, spacingAfter: null });
  expect(doc.paragraphs[0]!.directProperties().spacingBefore).toBe(31680); expect(doc.paragraphs[0]!.directProperties().spacingAfter).toBeNull();
  const before = doc.package.toBytes(); expect(() => doc.paragraphs[0]!.setProperties({ spacingBefore: 31681 })).toThrow(); expect(doc.package.toBytes()).toEqual(before);
  doc.paragraphs[0]!.setProperties({ spacingBefore: null });
  const spacing = elements(parseXml(new TextDecoder().decode(doc.package.get('word/document.xml'))), 'spacing', W)[0]!;
  expect(attribute(spacing, 'before', W)).toBeUndefined(); expect(attribute(spacing, 'after', W)).toBeUndefined();
  expect(attribute(spacing, 'line', W)).toBe('360'); expect(attribute(spacing, 'lineRule', W)).toBe('auto');
});

test('canonical flag predicates reject false readbacks rather than trusting setter success', async () => {
  const { fixturesRoot } = await import('../../scripts/fixture-inputs.ts'), { join } = await import('node:path');
  const { selectSharedScenarios, executeAcceptance } = await import('../../scripts/gherkin.ts');
  const { bindings } = await import('../acceptance/steps.ts');
  const path = 'workflows/docx/document-model.feature', text = await Bun.file(join(fixturesRoot(), path)).text();
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  const corrupt = bindings.map(b => b.pattern.test('KeepLines, PageBreakBefore and WidowControl are set true') ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => {
    await b.run(c, ...captures); (c.state as { document: Document }).document.paragraphs[0]!.setProperties({ widowControl: false });
  } } : b);
  const r = await executeAcceptance({ root: '.', features: [selectSharedScenarios(path, text, ['@id-docx-go-paragraph-advanced-toggles'])], counts: { features: count(1), scenarios: count(1), cases: count(1), steps: count(3) } }, corrupt, 'paragraph-flags-control');
  expect(r.counts.cases.failed).toBe(1); expect(r.counts.steps.failed).toBe(1); expect(r.counts.steps.undefined).toBe(0); expect(r.counts.steps.ambiguous).toBe(0);
});
