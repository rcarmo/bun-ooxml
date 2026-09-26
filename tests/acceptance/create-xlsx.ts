import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { elements, parseXml, type XmlElement } from "../../src/xml/index.ts";
import { Workbook } from "../../src/xlsx/index.ts";

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const encoder = new TextEncoder();
const projectRoot = join(import.meta.dir, "..", "..");

type ScenarioState = {
  sourceBytes?: Uint8Array;
  workbook?: Workbook;
  reopened?: Workbook;
  savedBytes?: Uint8Array;
  savedParts?: Map<string, Uint8Array>;
  originalParts?: Map<string, Uint8Array>;
  refusalCodes?: string[];
  appendOriginalSheetXml?: string;
};

export function createPrefixedWorkbookWithMissingB2(): Uint8Array {
  return writeZip(new Map<string, Uint8Array>([
    ["[Content_Types].xml", xml(`
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      </Types>
    `)],
    ["_rels/.rels", xml(`
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>
    `)],
    ["xl/workbook.xml", xml(`
      <x:workbook xmlns:x="${S_NS}" xmlns:link="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <x:sheets>
          <x:sheet name="Prefixed" sheetId="1" link:id="rId1"/>
        </x:sheets>
      </x:workbook>
    `)],
    ["xl/_rels/workbook.xml.rels", xml(`
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>
    `)],
    ["xl/styles.xml", xml(`
      <x:styleSheet xmlns:x="${S_NS}">
        <x:fonts count="1"><x:font/></x:fonts>
        <x:fills count="1"><x:fill/></x:fills>
        <x:borders count="1"><x:border/></x:borders>
        <x:cellStyleXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></x:cellStyleXfs>
        <x:cellXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></x:cellXfs>
        <x:cellStyles count="1"><x:cellStyle name="Normal" xfId="0" builtinId="0"/></x:cellStyles>
      </x:styleSheet>
    `)],
    ["xl/worksheets/sheet1.xml", xml(`
      <x:worksheet xmlns:x="${S_NS}">
        <x:sheetData>
          <x:row r="2"/>
        </x:sheetData>
      </x:worksheet>
    `)],
  ]));
}

export function createUnsafeWorksheetWorkbook(): Uint8Array {
  return writeZip(new Map<string, Uint8Array>([
    ["[Content_Types].xml", xml(`
      <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
      </Types>
    `)],
    ["_rels/.rels", xml(`
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
      </Relationships>
    `)],
    ["xl/workbook.xml", xml(`
      <workbook xmlns="${S_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets><sheet name="Unsafe" sheetId="1" r:id="rId1"/></sheets>
      </workbook>
    `)],
    ["xl/_rels/workbook.xml.rels", xml(`
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>
    `)],
    ["xl/styles.xml", xml(`
      <styleSheet xmlns="${S_NS}">
        <fonts count="1"><font/></fonts>
        <fills count="1"><fill/></fills>
        <borders count="1"><border/></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
      </styleSheet>
    `)],
    ["xl/worksheets/sheet1.xml", xml(`
      <worksheet xmlns="${S_NS}">
        <sheetData>
          <row r="2"><c r="A2"><v>2</v></c></row>
          <row r="1"><c r="A1"><v>1</v></c></row>
        </sheetData>
      </worksheet>
    `)],
  ]));
}

export const bindings: StepBinding[] = [
  {
    pattern: /^a native-created XLSX workbook$/,
    run: (context) => {
      const state = scenarioState(context);
      state.workbook = Workbook.create();
      state.sourceBytes = state.workbook.toBytes();
    },
  },
  {
    pattern: /^a prefixed XLSX fixture with a missing B2 target$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.sourceBytes = createPrefixedWorkbookWithMissingB2();
      state.workbook = await Workbook.open(state.sourceBytes);
    },
  },
  {
    pattern: /^the go-ooxml formatting workbook fixture for append$/,
    run: async (context) => {
      const state = scenarioState(context);
      const path = join(projectRoot, "fixtures/go-ooxml/testdata/excel/formatting.xlsx");
      state.sourceBytes = new Uint8Array(await Bun.file(path).arrayBuffer());
      state.originalParts = readZip(state.sourceBytes);
      state.appendOriginalSheetXml = decode(required(state.originalParts.get("xl/worksheets/sheet1.xml"), "sheet1.xml"));
      state.workbook = await Workbook.open(path);
    },
  },
  {
    pattern: /^I write "([^"]+)" to A1, (\d+) to B2, (true|false) to C3 and save and reopen it$/,
    run: async (context, textValue, numeric, booleanValue) => {
      const workbook = required(scenarioState(context).workbook, "workbook");
      workbook.worksheet("Sheet1").setCellValue("A1", textValue);
      workbook.worksheet("Sheet1").setCellValue("B2", Number(numeric));
      workbook.worksheet("Sheet1").setCellValue("C3", booleanValue === "true");
      await saveAndReopen(scenarioState(context), workbook);
    },
  },
  {
    pattern: /^I add worksheets "([^"]+)" and "([^"]+)", write (\d+) to Data A1 and "([^"]+)" to Summary B2, and save and reopen the workbook$/,
    run: async (context, dataName, summaryName, numeric, textValue) => {
      const workbook = required(scenarioState(context).workbook, "workbook");
      workbook.addWorksheet(dataName).setCellValue("A1", Number(numeric));
      workbook.addWorksheet(summaryName).setCellValue("B2", textValue);
      await saveAndReopen(scenarioState(context), workbook);
    },
  },
  {
    pattern: /^I write "([^"]+)" to its missing B2 and save and reopen it$/,
    run: async (context, textValue) => {
      const workbook = required(scenarioState(context).workbook, "workbook");
      workbook.worksheet("Prefixed").setCellValue("B2", textValue);
      await saveAndReopen(scenarioState(context), workbook);
    },
  },
  {
    pattern: /^I write (\d+) to XFD1048576 and save and reopen it$/,
    run: async (context, numeric) => {
      const workbook = required(scenarioState(context).workbook, "workbook");
      workbook.worksheet("Sheet1").setCellValue("XFD1048576", Number(numeric));
      await saveAndReopen(scenarioState(context), workbook);
    },
  },
  {
    pattern: /^I write (\d+) to C3, (\d+) to A1, (\d+) to B1 and (\d+) to A2 and save and reopen it$/,
    run: async (context, c3, a1, b1, a2) => {
      const workbook = required(scenarioState(context).workbook, "workbook");
      const sheet = workbook.worksheet("Sheet1");
      sheet.setCellValue("C3", Number(c3));
      sheet.setCellValue("A1", Number(a1));
      sheet.setCellValue("B1", Number(b1));
      sheet.setCellValue("A2", Number(a2));
      await saveAndReopen(scenarioState(context), workbook);
    },
  },
  {
    pattern: /^I try invalid worksheet coordinates and names$/,
    run: (context) => {
      const state = scenarioState(context);
      const workbook = required(state.workbook, "workbook");
      const before = workbook.toBytes();
      const refusalCodes: string[] = [];
      for (const coordinate of ["A0", "XFE1", "A1048577"]) {
        try {
          workbook.worksheet("Sheet1").setCellValue(coordinate, 1);
        } catch (error) {
          refusalCodes.push(codeOf(error));
        }
        assert.deepEqual(workbook.toBytes(), before);
      }
      for (const name of ["", "Sheet1", "sheet1", "'Quoted", "Quoted'", "bad/name", "x".repeat(32)]) {
        try {
          workbook.addWorksheet(name);
        } catch (error) {
          refusalCodes.push(codeOf(error));
        }
        assert.deepEqual(workbook.toBytes(), before);
      }
      state.refusalCodes = refusalCodes;
    },
  },
  {
    pattern: /^I append the missing cell E7 text to "([^"]+)" and save and reopen the workbook$/,
    run: async (context, textValue) => {
      const workbook = required(scenarioState(context).workbook, "workbook");
      workbook.worksheet("Formatted").setCellValue("E7", textValue);
      await saveAndReopen(scenarioState(context), workbook);
    },
  },
  {
    pattern: /^the reopened created workbook has Sheet1 values at A1, B2 and C3$/,
    run: (context) => {
      const workbook = required(scenarioState(context).reopened, "reopened");
      const sheet = workbook.worksheet("Sheet1");
      assert.equal(sheet.getCell("A1")?.value, "Alpha");
      assert.equal(sheet.getCell("B2")?.value, 7);
      assert.equal(sheet.getCell("C3")?.value, true);
    },
  },
  {
    pattern: /^the saved created workbook contains only the minimal workbook, worksheet and styles parts$/,
    run: (context) => {
      const parts = readZip(required(scenarioState(context).savedBytes, "savedBytes"));
      assert.deepEqual([...parts.keys()].sort(), [
        "[Content_Types].xml",
        "_rels/.rels",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
      ]);
    },
  },
  {
    pattern: /^Sheet1 dimension becomes (A1:C3|XFD1048576:XFD1048576)$/,
    run: (context, expectedDimension) => {
      const parts = required(scenarioState(context).savedParts, "savedParts");
      const worksheetDocument = parseXml(decode(required(parts.get("xl/worksheets/sheet1.xml"), "sheet1.xml")));
      const dimension = onlySpreadsheetElement(worksheetDocument.root, "dimension");
      assert.equal(dimension.attributes.ref, expectedDimension);
    },
  },
  {
    pattern: /^the reopened workbook sheetnames are Sheet1, Data and Summary$/,
    run: (context) => {
      assert.deepEqual(required(scenarioState(context).reopened, "reopened").sheetnames, ["Sheet1", "Data", "Summary"]);
      assert.equal(required(scenarioState(context).reopened, "reopened").worksheet("Data").getCell("A1")?.value, 5);
      assert.equal(required(scenarioState(context).reopened, "reopened").worksheet("Summary").getCell("B2")?.value, "done");
    },
  },
  {
    pattern: /^the saved workbook assigns independent worksheet relationship and sheet ids$/,
    run: (context) => {
      const parts = required(scenarioState(context).savedParts, "savedParts");
      const workbookDocument = parseXml(decode(required(parts.get("xl/workbook.xml"), "workbook.xml")));
      const relationshipsDocument = parseXml(decode(required(parts.get("xl/_rels/workbook.xml.rels"), "workbook.xml.rels")));
      const sheets = workbookDocument.root.children.find((child) => child.localName === "sheets" && child.namespaceURI === S_NS);
      assert.ok(sheets, "missing sheets element");
      const sheetElements = sheets.children.filter((child) => child.localName === "sheet" && child.namespaceURI === S_NS);
      assert.deepEqual(sheetElements.map((sheet) => sheet.attributes.name), ["Sheet1", "Data", "Summary"]);
      assert.deepEqual(sheetElements.map((sheet) => sheet.attributes.sheetId), ["1", "2", "3"]);
      const relationshipIds = sheetElements.map((sheet) => relationshipId(sheet));
      assert.deepEqual(relationshipIds, ["rId1", "rId3", "rId4"]);
      const worksheetRelationships = relationshipsDocument.root.children.filter((child) =>
        child.localName === "Relationship"
        && child.attributes.Type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      );
      assert.deepEqual(worksheetRelationships.map((relationship) => relationship.attributes.Id), ["rId1", "rId3", "rId4"]);
      assert.deepEqual(worksheetRelationships.map((relationship) => relationship.attributes.Target), [
        "worksheets/sheet1.xml",
        "worksheets/sheet2.xml",
        "worksheets/sheet3.xml",
      ]);
    },
  },
  {
    pattern: /^the saved prefixed worksheet uses qualified dimension row cell and text nodes for B2$/,
    run: (context) => {
      const parts = required(scenarioState(context).savedParts, "savedParts");
      const worksheetDocument = parseXml(decode(required(parts.get("xl/worksheets/sheet1.xml"), "sheet1.xml")));
      const dimension = onlySpreadsheetElement(worksheetDocument.root, "dimension");
      const row = findRow(worksheetDocument.root, 2);
      const cell = findCell(worksheetDocument.root, "B2");
      const inline = required(directSpreadsheetChild(cell, "is"), "inline string");
      const text = required(directSpreadsheetChild(inline, "t"), "inline text");
      assert.equal(dimension.name, "x:dimension");
      assert.equal(dimension.attributes.ref, "B2:B2");
      assert.equal(row.name, "x:row");
      assert.equal(cell.name, "x:c");
      assert.equal(inline.name, "x:is");
      assert.equal(text.name, "x:t");
      assert.equal(text.text, "Prefixed");
    },
  },
  {
    pattern: /^the reopened prefixed workbook reads B2 as "([^"]+)"$/,
    run: (context, expectedValue) => {
      assert.equal(required(scenarioState(context).reopened, "reopened").worksheet("Prefixed").getCell("B2")?.value, expectedValue);
    },
  },
  {
    pattern: /^the reopened created workbook reads XFD1048576 as 1$/,
    run: (context) => {
      assert.equal(required(scenarioState(context).reopened, "reopened").worksheet("Sheet1").getCell("XFD1048576")?.value, 1);
    },
  },
  {
    pattern: /^the saved created worksheet stores rows and cells in ascending numeric order$/,
    run: (context) => {
      const worksheetDocument = parseXml(decode(required(required(scenarioState(context).savedParts, "savedParts").get("xl/worksheets/sheet1.xml"), "sheet1.xml")));
      const sheetData = onlySpreadsheetElement(worksheetDocument.root, "sheetData");
      const rows = sheetData.children.filter((child) => child.localName === "row" && child.namespaceURI === S_NS);
      assert.deepEqual(rows.map((row) => row.attributes.r), ["1", "2", "3"]);
      assert.deepEqual(rows.map((row) => directSpreadsheetChildren(row, "c").map((cell) => cell.attributes.r)), [
        ["A1", "B1"],
        ["A2"],
        ["C3"],
      ]);
    },
  },
  {
    pattern: /^the reopened created workbook keeps those ordered values$/,
    run: (context) => {
      const sheet = required(scenarioState(context).reopened, "reopened").worksheet("Sheet1");
      assert.equal(sheet.getCell("A1")?.value, 1);
      assert.equal(sheet.getCell("B1")?.value, 2);
      assert.equal(sheet.getCell("A2")?.value, 4);
      assert.equal(sheet.getCell("C3")?.value, 3);
    },
  },
  {
    pattern: /^every invalid create-side refusal leaves the workbook bytes unchanged$/,
    run: (context) => {
      assert.deepEqual(required(scenarioState(context).refusalCodes, "refusalCodes"), [
        "xlsx-cell-reference-invalid",
        "xlsx-cell-reference-invalid",
        "xlsx-cell-reference-invalid",
        "xlsx-worksheet-name-invalid",
        "xlsx-worksheet-duplicate",
        "xlsx-worksheet-duplicate",
        "xlsx-worksheet-name-invalid",
        "xlsx-worksheet-name-invalid",
        "xlsx-worksheet-name-invalid",
        "xlsx-worksheet-name-invalid",
      ]);
    },
  },
  {
    pattern: /^the workbook still has only Sheet1$/,
    run: (context) => {
      assert.deepEqual(required(scenarioState(context).workbook, "workbook").sheetnames, ["Sheet1"]);
    },
  },
  {
    pattern: /^the reopened existing worksheet has the appended value at E7$/,
    run: (context) => {
      const cell = required(scenarioState(context).reopened, "reopened").worksheet("Formatted").getCell("E7");
      assert.equal(cell?.value, "tail");
      assert.equal(cell?.styleId, undefined);
    },
  },
  {
    pattern: /^the existing worksheet keeps its surrounding nodes and unrelated parts$/,
    run: (context) => {
      const state = scenarioState(context);
      const originalParts = required(state.originalParts, "originalParts");
      const savedParts = required(state.savedParts, "savedParts");
      for (const name of ["[Content_Types].xml", "docProps/core.xml", "xl/styles.xml", "xl/theme/theme1.xml", "xl/workbook.xml"]) {
        equalBytes(originalParts.get(name), savedParts.get(name), name);
      }
      const originalSheetXml = required(state.appendOriginalSheetXml, "appendOriginalSheetXml");
      const savedSheetXml = decode(required(savedParts.get("xl/worksheets/sheet1.xml"), "sheet1.xml"));
      assert.match(savedSheetXml, /<cols><col width="20" customWidth="1" min="1" max="1"\/><col width="20" customWidth="1" min="2" max="2"\/><col width="20" customWidth="1" min="3" max="3"\/><col width="20" customWidth="1" min="4" max="4"\/><\/cols>/);
      assert.match(savedSheetXml, /<pageMargins left="0\.75" right="0\.75" top="1" bottom="1" header="0\.5" footer="0\.5"\/>/);
      assert.ok(originalSheetXml.includes("<sheetData>"));
      assert.match(savedSheetXml, /<dimension ref="A1:E7"\/>/);
    },
  },
];

export default bindings;

function scenarioState(context: Record<string, unknown>): ScenarioState {
  return context.state as ScenarioState;
}

async function saveAndReopen(state: ScenarioState, workbook: Workbook): Promise<void> {
  const outputPath = await tempWorkbookPath("create-xlsx.xlsx");
  await workbook.save(outputPath);
  state.savedBytes = new Uint8Array(await Bun.file(outputPath).arrayBuffer());
  state.savedParts = readZip(state.savedBytes);
  state.reopened = await Workbook.open(state.savedBytes);
}

function required<T>(value: T | undefined, label: string): T {
  assert.ok(value !== undefined, `missing ${label}`);
  return value;
}

function codeOf(error: unknown): string {
  assert.ok(error && typeof error === "object" && "code" in error, "missing refusal code");
  return String((error as { code: unknown }).code);
}

function decode(bytes: Uint8Array): string {
  return new TextDecoder().decode(bytes);
}

function equalBytes(left: Uint8Array | undefined, right: Uint8Array | undefined, label: string): void {
  const leftBytes = required(left, `${label} left`);
  const rightBytes = required(right, `${label} right`);
  assert.equal(Buffer.compare(Buffer.from(leftBytes), Buffer.from(rightBytes)), 0, `${label} bytes differ`);
}

function onlySpreadsheetElement(root: XmlElement, localName: string): XmlElement {
  const matches = elements(root, localName, S_NS);
  assert.equal(matches.length, 1, `expected exactly one ${localName}`);
  return required(matches[0], localName);
}

function directSpreadsheetChild(element: XmlElement, localName: string): XmlElement | undefined {
  return element.children.find((child) => child.localName === localName && child.namespaceURI === S_NS);
}

function directSpreadsheetChildren(element: XmlElement, localName: string): XmlElement[] {
  return element.children.filter((child) => child.localName === localName && child.namespaceURI === S_NS);
}

function findCell(root: XmlElement, ref: string): XmlElement {
  const cell = elements(root, "c", S_NS).find((element) => element.attributes.r === ref);
  assert.ok(cell, `missing cell ${ref}`);
  return cell;
}

function findRow(root: XmlElement, rowNumber: number): XmlElement {
  const row = elements(root, "row", S_NS).find((element) => element.attributes.r === String(rowNumber));
  assert.ok(row, `missing row ${rowNumber}`);
  return row;
}

function relationshipId(sheet: XmlElement): string {
  const match = Object.keys(sheet.attributes).find((name) => {
    const localName = name.includes(":") ? name.slice(name.indexOf(":") + 1) : name;
    return localName === "id" && sheet.attributeNamespaces[name] === "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  });
  assert.ok(match, "missing relationship id");
  return sheet.attributes[match]!;
}

function xml(source: string): Uint8Array {
  const body = source.replace(/^\s+|\s+$/g, "").replace(/>\s+</g, "><");
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>${body}`);
}

async function tempWorkbookPath(name: string): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-create-xlsx-"));
  return join(root, name);
}
