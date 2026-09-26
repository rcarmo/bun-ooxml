import {fixturePath,F} from "../../scripts/fixture-inputs.ts";
import { afterEach, describe, expect, test } from "bun:test";
import {mergedNestedTableDocument} from '../fixtures/native-edge-cases.ts';
import { readFileSync } from "node:fs";
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { bindings as tableBindings } from "../acceptance/tables-docx.ts";
import { Document, W_NS, type TableCell } from "../../src/docx/index.ts";
import { OoxmlError } from "../../src/errors.ts";
import { OpcPackage } from '../../src/opc/package.ts';
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { applyEdits, attribute, elements, parseXml, type XmlElement } from "../../src/xml/index.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const DOCUMENT_PART = "word/document.xml";
const UTF8_DECODER = new TextDecoder("utf-8", { fatal: true });
const UTF8_ENCODER = new TextEncoder();
const tempRoots: string[] = [];

afterEach(async () => {
  await Promise.all(tempRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("docx Document tables", () => {
  test('authoring into a body with w rebound never silently writes foreign elements',async()=>{
    const created=Document.create();created.addParagraph('source');const p=await OpcPackage.open(await created.save());
    let xml=p.text('word/document.xml').replaceAll('xmlns:w=','xmlns:z=').replaceAll('w:','z:');xml=xml.replace('<z:body>', '<z:body xmlns:w="urn:foreign">');p.set('word/document.xml',xml);
    const doc=await Document.open(p.toBytes());const table=doc.addTable(1,1);table.cell(0,0).text='correct';doc.addParagraph('after');
    const after=await Document.open(await doc.save());expect(after.tables.length).toBe(1);expect(after.tables[0]!.cell(0,0).text).toBe('correct');expect(after.paragraphs.some(p=>p.text==='after')).toBe(true);
  });
  test('run-level namespace declarations survive copied table-cell run properties',async()=>{
    const doc=Document.create();doc.addTable(1,1).cell(0,0).text='before';const p=await OpcPackage.open(await doc.save());
    p.set('word/document.xml',p.text('word/document.xml').replace(/<w:r(?:\s[^>]*)?><w:t>before<\/w:t><\/w:r>/,'<w:r xmlns:f="urn:format"><w:rPr><f:flag/></w:rPr><w:t>before</w:t></w:r>'));
    const loaded=await Document.open(p.toBytes());loaded.tables[0]!.cell(0,0).text='after';const xml=parseXml(new TextDecoder().decode(loaded.package.get('word/document.xml')!));
    expect(elements(xml,'flag','urn:format').length).toBe(1);expect(loaded.tables[0]!.cell(0,0).text).toBe('after');
  });
  test("adds a rectangular table before sectPr and round-trips escaped text, spaces, and normalized newlines", async () => {
    const doc = Document.create();

    const table = doc.addTable(2, 2);
    table.cell(0, 0).text = "  <Alpha & Beta>  ";
    table.cell(1, 1).text = "line 1\r\nline 2";

    const reopened = await Document.open(await doc.save());
    const reopenedTable = requireTable(reopened, 0);

    expect(table.rows).toBe(2);
    expect(table.columns).toBe(2);
    expect(reopened.tables).toHaveLength(1);
    expect(reopenedTable.rows).toBe(2);
    expect(reopenedTable.columns).toBe(2);
    expect(reopenedTable.cell(0, 0).text).toBe("  <Alpha & Beta>  ");
    expect(reopenedTable.cell(1, 1).text).toBe("line 1\nline 2");

    const bodyChildren = readBody(reopened).children.filter((child) => isWord(child, "tbl") || isWord(child, "sectPr"));
    expect(bodyChildren.map((child) => child.localName)).toEqual(["tbl", "sectPr"]);

    const firstCellXml = rawTopLevelCellXml(reopened, 0, 0, 0);
    expect(firstCellXml).toContain('<w:t xml:space="preserve">  &lt;Alpha &amp; Beta&gt;  </w:t>');

    const secondCell = readTopLevelCellElement(reopened, 0, 1, 1);
    expect(secondCell.children.filter((child) => isWord(child, "p"))).toHaveLength(2);
  });

  test("updates existing simple table cells while preserving opaque package parts", async () => {
    const doc = await Document.open(fixturePath(F.officeWord.simpleTable));
    const originalCore = requireDefined(doc.package.get("docProps/core.xml"), "Missing docProps/core.xml");

    const table = requireTable(doc, 0);
    table.cell(1, 0).text = "Voltaic battery";

    const reopened = await Document.open(await doc.save());
    expect(requireTable(reopened, 0).cell(1, 0).text).toBe("Voltaic battery");
    expect(doc.package.get("docProps/core.xml")).toEqual(originalCore);
    expect(doc.package.diff()).toEqual({
      added: [],
      changed: [DOCUMENT_PART],
      removed: [],
    });
  });

  test("preserves tcPr, first paragraph properties, and first run formatting when replacing a simple cell", async () => {
    const doc = await Document.open(buildStyledCellFixture());

    const cell = requireTable(doc, 0).cell(1, 0);
    cell.text = "updated value";

    const reopened = await Document.open(await doc.save());
    const reopenedCell = readTopLevelCellElement(reopened, 0, 1, 0);
    const tcPr = requireDefined(reopenedCell.children.find((child) => isWord(child, "tcPr")), "Missing w:tcPr");
    const shading = requireDefined(tcPr.children.find((child) => isWord(child, "shd")), "Missing w:shd");
    expect(attribute(shading, "fill", W_NS)).toBe("FFFF00");

    const firstParagraph = requireDefined(reopenedCell.children.find((child) => isWord(child, "p")), "Missing first w:p");
    const paragraphProperties = requireDefined(firstParagraph.children.find((child) => isWord(child, "pPr")), "Missing w:pPr");
    const justification = requireDefined(paragraphProperties.children.find((child) => isWord(child, "jc")), "Missing w:jc");
    expect(attribute(justification, "val", W_NS)).toBe("right");
    const run = requireDefined(firstParagraph.children.find((child) => isWord(child, "r")), "Missing first w:r");
    const runProperties = requireDefined(run.children.find((child) => isWord(child, "rPr")), "Missing w:rPr");
    expect(runProperties.children.some((child) => isWord(child, "b"))).toBeTrue();
    expect(requireTable(reopened, 0).cell(1, 0).text).toBe("updated value");
  });

  test("refuses bad indices, merged cells, nested cell topology, and bizarre grids atomically", async () => {
    const simple = await Document.open(fixturePath(F.officeWord.simpleTable));
    const simpleBaseline = await simple.save();
    expect(() => requireTable(simple, 0).cell(9, 0)).toThrow(RangeError);
    expect(await simple.save()).toEqual(simpleBaseline);

    const merged = await Document.open(await mergedNestedTableDocument());
    const mergedBaseline = await merged.save();
    const mergedTable = requireTable(merged, 0);
    expect(mergedTable.rows).toBe(3);
    expect(mergedTable.columns).toBe(3);
    expect(() => mergedTable.cell(0, 0)).toThrow(expect.objectContaining({ code: "docx-table-merged-cell" }));
    expect(await merged.save()).toEqual(mergedBaseline);

    const nestedCell = mergedTable.cell(2, 0);
    expect(() => {
      nestedCell.text = "updated";
    }).toThrow(expect.objectContaining({ code: "docx-table-cell-unsupported" }));
    expect(await merged.save()).toEqual(mergedBaseline);

    const bizarre = await Document.open(buildGridBeforeFixture());
    const bizarreBaseline = await bizarre.save();
    expect(() => requireTable(bizarre, 0).cell(0, 0)).toThrow(
      expect.objectContaining({ code: "docx-table-unsupported" }),
    );
    expect(await bizarre.save()).toEqual(bizarreBaseline);
  });

  test("stale table and cell handles refuse after document mutation while fresh handles stay live", async () => {
    const doc = Document.create();
    const table = doc.addTable(1, 1);
    const staleCell = table.cell(0, 0);

    doc.addParagraph("after table");
    const bytesAfterEdit = await doc.save();

    expect(() => table.rows).toThrow(expect.objectContaining({ code: "docx-stale-table" }));
    expect(() => {
      staleCell.text = "stale write";
    }).toThrow(expect.objectContaining({ code: "docx-stale-table-cell" }));
    expect(await doc.save()).toEqual(bytesAfterEdit);

    const freshCell = requireTable(doc, 0).cell(0, 0);
    freshCell.text = "live write";
    const reopened = await Document.open(await doc.save());
    expect(requireTable(reopened, 0).cell(0, 0).text).toBe("live write");
  });

  test("validates table dimensions with bounded rectangular authoring", async () => {
    const doc = Document.create();
    const baseline = await doc.save();

    expect(() => doc.addTable(0, 1)).toThrow(RangeError);
    expect(() => doc.addTable(1, 0)).toThrow(RangeError);
    expect(() => doc.addTable(101, 100)).toThrow(RangeError);
    expect(await doc.save()).toEqual(baseline);
  });

  test("passes the DOCX table acceptance feature with dedicated bindings", async () => {
    const root = await makeAcceptanceRoot();
    const report = await runAcceptance(tableBindings, { root });

    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.inventory.scenarios.implemented).toBe(4);
    expect(report.execution.cases.passed).toBe(7);
    expect(report.execution.steps.failed).toBe(0);
    expect(report.failures).toEqual([]);
  }, 300_000);
});

function buildStyledCellFixture(): Uint8Array {
  const source = new Uint8Array(readFileSync(fixturePath(F.officeWord.simpleTable)));
  const parts = readZip(source);
  const xml = UTF8_DECODER.decode(requireDefined(parts.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`));
  const cell = findTopLevelCell(parseXml(xml), 0, 1, 0);
  const updated = applyEdits(xml, [
    {
      start: cell.openEnd,
      end: cell.closeStart,
      value: [
        '<w:tcPr><w:tcW w:w="2880" w:type="dxa"/><w:shd w:fill="FFFF00"/></w:tcPr>',
        '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Galvanic battery</w:t></w:r></w:p>',
      ].join(""),
    },
  ]);
  parts.set(DOCUMENT_PART, UTF8_ENCODER.encode(updated));
  return writeZip(parts);
}

function buildGridBeforeFixture(): Uint8Array {
  const source = new Uint8Array(readFileSync(fixturePath(F.officeWord.simpleTable)));
  const parts = readZip(source);
  const xml = UTF8_DECODER.decode(requireDefined(parts.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`));
  const parsed = parseXml(xml);
  const tableElement = requireDefined(readBodyElement(parsed).children.find((child) => isWord(child, "tbl")), "Missing first w:tbl");
  const firstRow = requireDefined(tableElement.children.find((child) => isWord(child, "tr")), "Missing first w:tr");
  const firstCell = requireDefined(firstRow.children.find((child) => isWord(child, "tc")), "Missing first w:tc");
  const updated = applyEdits(xml, [
    { start: firstCell.start, end: firstCell.end, value: "" },
    { start: firstRow.openEnd, end: firstRow.openEnd, value: '<w:trPr><w:gridBefore w:val="1"/></w:trPr>' },
  ]);
  parts.set(DOCUMENT_PART, UTF8_ENCODER.encode(updated));
  return writeZip(parts);
}

function findTopLevelCell(parsed: ReturnType<typeof parseXml>, tableIndex: number, row: number, column: number): XmlElement {
  const tableElement = requireDefined(
    readBodyElement(parsed).children.filter((child) => isWord(child, "tbl"))[tableIndex],
    `Missing DOCX table ${tableIndex + 1}`,
  );
  const rowElement = requireDefined(
    tableElement.children.filter((child) => isWord(child, "tr"))[row],
    `Missing DOCX row ${row}`,
  );
  return requireDefined(
    rowElement.children.filter((child) => isWord(child, "tc"))[column],
    `Missing DOCX cell (${row},${column})`,
  );
}

function rawTopLevelCellXml(doc: Document, tableIndex: number, row: number, column: number): string {
  const cell = readTopLevelCellElement(doc, tableIndex, row, column);
  const xml = readDocumentXml(doc);
  return xml.slice(cell.start, cell.end);
}

function readTopLevelCellElement(doc: Document, tableIndex: number, row: number, column: number): XmlElement {
  return findTopLevelCell(parseXml(readDocumentXml(doc)), tableIndex, row, column);
}

function readBody(doc: Document): XmlElement {
  return readBodyElement(parseXml(readDocumentXml(doc)));
}

function readBodyElement(parsed: ReturnType<typeof parseXml>): XmlElement {
  return requireDefined(parsed.root.children.find((child) => isWord(child, "body")), "Missing w:body");
}

function readDocumentXml(doc: Document): string {
  return UTF8_DECODER.decode(requireDefined(doc.package.get(DOCUMENT_PART), `Missing ${DOCUMENT_PART}`));
}

function requireTable(doc: Document, index: number) {
  return requireDefined(doc.tables[index], `Missing DOCX table ${index + 1}`);
}

function requireDefined<T>(value: T | undefined, message: string): T {
  expect(value).toBeDefined();
  if (value === undefined) {
    throw new Error(message);
  }
  return value;
}

function isWord(element: XmlElement, localName: string): boolean {
  return element.localName === localName && element.namespaceURI === W_NS;
}

async function makeAcceptanceRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-docx-table-acceptance-"));
  tempRoots.push(root);
  const featureText = await Bun.file(join(PROJECT_ROOT, "features/docx/tables.feature")).text();
  const path = join(root, "features", "docx", "tables.feature");
  await mkdir(dirname(path), { recursive: true });
  await Bun.write(path, featureText);
  return root;
}
