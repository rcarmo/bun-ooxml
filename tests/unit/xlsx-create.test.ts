import { afterEach, describe, expect, test } from "bun:test";
import { namespaceFixture } from '../acceptance/relationship-namespaces.ts';
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip } from "../../src/opc/zip.ts";
import { elements, parseXml, type XmlElement } from "../../src/xml/index.ts";
import { Workbook } from "../../src/xlsx/index.ts";
import { fixturePath, F } from "../../scripts/fixture-inputs.ts";
import {
  bindings,
  createPrefixedWorkbookWithMissingB2,
  createUnsafeWorksheetWorkbook,
} from "../acceptance/create-xlsx.ts";

test('adding a worksheet works when the existing r alias is scoped to only the first sheet',async()=>{
 const workbook=await Workbook.open(await namespaceFixture('xlsx','alias'));const before=workbook.package.get('xl/worksheets/sheet1.xml')!;
 const added=workbook.addWorksheet('Second');added.setCellValue('A1','new');const reopened=await Workbook.open(workbook.toBytes());
 expect(reopened.sheetnames).toEqual(['Sheet','Second']);expect(reopened.worksheet('Second').getCell('A1')?.value).toBe('new');expect(reopened.package.get('xl/worksheets/sheet1.xml')).toEqual(before);
});

test('dimension stays before sheetViews and cols when a missing cell is authored',async()=>{
 const w=Workbook.create(),part='xl/worksheets/sheet1.xml';
 let xml=w.package.text(part).replace(/<dimension[^>]*\/>/,'');xml=xml.replace('<sheetData', '<sheetViews><sheetView workbookViewId="0"/></sheetViews><cols><col min="1" max="1" width="12"/></cols><sheetData');w.package.set(part,xml);
 const loaded=await Workbook.open(w.toBytes());loaded.worksheet('Sheet1').setCellValue('C3',7);
 const names=parseXml(loaded.package.text(part)).root.children.map(n=>n.localName);
 expect(names.indexOf('dimension')).toBeLessThan(names.indexOf('sheetViews'));expect(names.indexOf('dimension')).toBeLessThan(names.indexOf('cols'));
});

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const roots: string[] = [];

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("Workbook.create", () => {
  test("authors a minimal workbook with styles that saves and reopens through the native reader", async () => {
    const workbook = Workbook.create();
    const sheet = workbook.worksheet("Sheet1");

    sheet.setCellValue("A1", "Alpha");
    sheet.setCellValue("B2", 7);
    sheet.setCellValue("C3", true);

    const outPath = join(await tempRoot(), "created.xlsx");
    await workbook.save(outPath);
    const savedBytes = new Uint8Array(await Bun.file(outPath).arrayBuffer());
    const savedParts = readZip(savedBytes);
    const reopened = await Workbook.open(savedBytes);
    const worksheetDocument = parseXml(decodePart(savedParts, "xl/worksheets/sheet1.xml"));
    const dimension = onlySpreadsheetElement(worksheetDocument.root, "dimension");

    expect(workbook.sheetnames).toEqual(["Sheet1"]);
    expect(reopened.sheetnames).toEqual(["Sheet1"]);
    expect(reopened.worksheet("Sheet1").getCell("A1")).toMatchObject({ kind: "string", value: "Alpha" });
    expect(reopened.worksheet("Sheet1").getCell("B2")).toMatchObject({ kind: "number", value: 7 });
    expect(reopened.worksheet("Sheet1").getCell("C3")).toMatchObject({ kind: "boolean", value: true });
    expect(dimension.attributes.ref).toBe("A1:C3");
    expect([...savedParts.keys()].sort()).toEqual([
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/_rels/workbook.xml.rels",
      "xl/styles.xml",
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
    ]);
    expect(savedParts.get("xl/styles.xml")).toBeDefined();
  });

  test("adds worksheets with independent ids, relationship closure and working returned handles", async () => {
    const workbook = Workbook.create();
    const data = workbook.addWorksheet("Data");
    const summary = workbook.addWorksheet("Summary");

    data.setCellValue("A1", 5);
    summary.setCellValue("B2", "done");

    const bytes = workbook.toBytes();
    const parts = readZip(bytes);
    const workbookDocument = parseXml(decodePart(parts, "xl/workbook.xml"));
    const workbookRelationships = parseXml(decodePart(parts, "xl/_rels/workbook.xml.rels"));
    const reopened = await Workbook.open(bytes);
    const sheets = onlySpreadsheetElement(workbookDocument.root, "sheets").children
      .filter((child) => child.localName === "sheet" && child.namespaceURI === S_NS);
    const worksheetRelationships = workbookRelationships.root.children.filter((child) =>
      child.localName === "Relationship"
      && child.attributes.Type === "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    );

    expect(reopened.sheetnames).toEqual(["Sheet1", "Data", "Summary"]);
    expect(reopened.worksheet("Data").getCell("A1")?.value).toBe(5);
    expect(reopened.worksheet("Summary").getCell("B2")?.value).toBe("done");
    expect(sheets.map((sheet) => sheet.attributes.name)).toEqual(["Sheet1", "Data", "Summary"]);
    expect(sheets.map((sheet) => sheet.attributes.sheetId)).toEqual(["1", "2", "3"]);
    expect(sheets.map((sheet) => relationshipId(sheet))).toEqual(["rId1", "rId3", "rId4"]);
    expect(worksheetRelationships.map((relationship) => relationship.attributes.Id)).toEqual(["rId1", "rId3", "rId4"]);
    expect(worksheetRelationships.map((relationship) => relationship.attributes.Target)).toEqual([
      "worksheets/sheet1.xml",
      "worksheets/sheet2.xml",
      "worksheets/sheet3.xml",
    ]);
    expect(workbook.package.diff()).toEqual({
      added: ["xl/worksheets/sheet2.xml", "xl/worksheets/sheet3.xml"],
      changed: ["[Content_Types].xml", "xl/_rels/workbook.xml.rels", "xl/workbook.xml"],
      removed: [],
    });
  });

  test("refuses invalid worksheet names atomically with case-insensitive duplicate checks", () => {
    const workbook = Workbook.create();
    const before = workbook.toBytes();

    const cases = [
      ["", "xlsx-worksheet-name-invalid"],
      ["Sheet1", "xlsx-worksheet-duplicate"],
      ["sheet1", "xlsx-worksheet-duplicate"],
      ["'Quoted", "xlsx-worksheet-name-invalid"],
      ["Quoted'", "xlsx-worksheet-name-invalid"],
      ["bad/name", "xlsx-worksheet-name-invalid"],
      ["x".repeat(32), "xlsx-worksheet-name-invalid"],
    ] as const;

    for (const [name, code] of cases) {
      expect(() => workbook.addWorksheet(name)).toThrow(
        expect.objectContaining({ code } satisfies Partial<OoxmlError>),
      );
      expect(workbook.toBytes()).toEqual(before);
      expect(workbook.sheetnames).toEqual(["Sheet1"]);
    }
  });
});

describe("Worksheet.setCellValue missing-cell authoring", () => {
  test("creates qualified dimension row cell and text nodes in a prefixed worksheet", async () => {
    const workbook = await Workbook.open(createPrefixedWorkbookWithMissingB2());

    workbook.worksheet("Prefixed").setCellValue("B2", "Prefixed");

    const bytes = workbook.toBytes();
    const parts = readZip(bytes);
    const worksheetDocument = parseXml(decodePart(parts, "xl/worksheets/sheet1.xml"));
    const dimension = onlySpreadsheetElement(worksheetDocument.root, "dimension");
    const row = findRow(worksheetDocument.root, 2);
    const cell = findCell(worksheetDocument.root, "B2");
    const inline = onlyDirectSpreadsheetChild(cell, "is");
    const text = onlyDirectSpreadsheetChild(inline, "t");
    const reopened = await Workbook.open(bytes);

    expect(dimension.name).toBe("x:dimension");
    expect(dimension.attributes.ref).toBe("B2:B2");
    expect(row.name).toBe("x:row");
    expect(cell.name).toBe("x:c");
    expect(inline.name).toBe("x:is");
    expect(text.name).toBe("x:t");
    expect(text.text).toBe("Prefixed");
    expect(reopened.worksheet("Prefixed").getCell("B2")?.value).toBe("Prefixed");
  });

  test("accepts the maximum worksheet coordinate and keeps rows and cells sorted as they are added", async () => {
    const workbook = Workbook.create();
    const sheet = workbook.worksheet("Sheet1");

    sheet.setCellValue("XFD1048576", 1);
    sheet.setCellValue("C3", 3);
    sheet.setCellValue("A1", 1);
    sheet.setCellValue("B1", 2);
    sheet.setCellValue("A2", 4);

    const bytes = workbook.toBytes();
    const parts = readZip(bytes);
    const worksheetDocument = parseXml(decodePart(parts, "xl/worksheets/sheet1.xml"));
    const sheetData = onlySpreadsheetElement(worksheetDocument.root, "sheetData");
    const rows = sheetData.children.filter((child) => child.localName === "row" && child.namespaceURI === S_NS);
    const reopened = await Workbook.open(bytes);

    expect(reopened.worksheet("Sheet1").getCell("XFD1048576")?.value).toBe(1);
    expect(onlySpreadsheetElement(worksheetDocument.root, "dimension").attributes.ref).toBe("A1:XFD1048576");
    expect(rows.map((row) => row.attributes.r)).toEqual(["1", "2", "3", "1048576"]);
    expect(rows.map((row) => directSpreadsheetChildren(row, "c").map((cell) => cell.attributes.r))).toEqual([
      ["A1", "B1"],
      ["A2"],
      ["C3"],
      ["XFD1048576"],
    ]);
  });

  test("refuses invalid coordinates without mutating the workbook", () => {
    const workbook = Workbook.create();
    const before = workbook.toBytes();

    for (const coordinate of ["A0", "XFE1", "A1048577"]) {
      expect(() => workbook.worksheet("Sheet1").setCellValue(coordinate, 1)).toThrow(
        expect.objectContaining({
          code: "xlsx-cell-reference-invalid",
        } satisfies Partial<OoxmlError>),
      );
      expect(workbook.toBytes()).toEqual(before);
      expect(workbook.package.diff()).toEqual({ added: [], changed: [], removed: [] });
    }
  });

  test("refuses unsafe existing row ordering before mutation", async () => {
    const originalBytes = createUnsafeWorksheetWorkbook();
    const workbook = await Workbook.open(originalBytes);

    expect(() => workbook.worksheet("Unsafe").setCellValue("B1", 9)).toThrow(
      expect.objectContaining({
        code: "xlsx-worksheet-structure-unsupported",
      } satisfies Partial<OoxmlError>),
    );
    expect(workbook.toBytes()).toEqual(originalBytes);
    expect(workbook.package.diff()).toEqual({ added: [], changed: [], removed: [] });
  });

  test("appends a missing cell to an existing worksheet while preserving surrounding nodes and unrelated parts", async () => {
    const path = fixturePath(F.goSheets.formatting);
    const originalBytes = new Uint8Array(await Bun.file(path).arrayBuffer());
    const originalParts = readZip(originalBytes);
    const originalSheetXml = decodePart(originalParts, "xl/worksheets/sheet1.xml");
    const workbook = await Workbook.open(path);

    workbook.worksheet("Formatted").setCellValue("E7", "tail");

    const outPath = join(await tempRoot(), "append-existing.xlsx");
    await workbook.save(outPath);
    const savedBytes = new Uint8Array(await Bun.file(outPath).arrayBuffer());
    const savedParts = readZip(savedBytes);
    const savedSheetXml = decodePart(savedParts, "xl/worksheets/sheet1.xml");
    const reopened = await Workbook.open(savedBytes);

    expect(reopened.worksheet("Formatted").getCell("E7")).toMatchObject({
      kind: "string",
      value: "tail",
      styleId: undefined,
    });
    expect(savedSheetXml).toContain("<cols><col width=\"20\" customWidth=\"1\" min=\"1\" max=\"1\"/><col width=\"20\" customWidth=\"1\" min=\"2\" max=\"2\"/><col width=\"20\" customWidth=\"1\" min=\"3\" max=\"3\"/><col width=\"20\" customWidth=\"1\" min=\"4\" max=\"4\"/></cols>");
    expect(savedSheetXml).toContain("<pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/>");
    expect(savedSheetXml).toContain("<dimension ref=\"A1:E7\"/>");
    expect(originalSheetXml).toContain("<sheetData>");
    expect(savedParts.get("[Content_Types].xml")).toEqual(originalParts.get("[Content_Types].xml"));
    expect(savedParts.get("docProps/core.xml")).toEqual(originalParts.get("docProps/core.xml"));
    expect(savedParts.get("xl/styles.xml")).toEqual(originalParts.get("xl/styles.xml"));
    expect(savedParts.get("xl/theme/theme1.xml")).toEqual(originalParts.get("xl/theme/theme1.xml"));
    expect(savedParts.get("xl/workbook.xml")).toEqual(originalParts.get("xl/workbook.xml"));
    expect(workbook.package.diff()).toEqual({
      added: [],
      changed: ["xl/worksheets/sheet1.xml"],
      removed: [],
    });
  });
});

describe("xlsx create acceptance feature", () => {
  test("executes the implemented xlsx creation scenarios with exported bindings", async () => {
    const root = await tempRoot();
    const featureText = await Bun.file(join(import.meta.dir, "../../features/xlsx/create.feature")).text();
    await mkdir(join(root, "features", "xlsx"), { recursive: true });
    await Bun.write(join(root, "features/xlsx/create.feature"), featureText);

    const report = await runAcceptance(bindings, { root });

    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.inventory.scenarios.implemented).toBe(7);
    expect(report.execution.cases.failed).toBe(0);
    expect(report.execution.steps.undefined).toBe(0);
  });
});

async function tempRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-xlsx-create-test-"));
  roots.push(root);
  return root;
}

function decodePart(parts: Map<string, Uint8Array>, name: string): string {
  return new TextDecoder().decode(parts.get(name)!);
}

function onlySpreadsheetElement(root: XmlElement, localName: string): XmlElement {
  const matches = elements(root, localName, S_NS);
  expect(matches.length).toBe(1);
  return matches[0]!;
}

function findCell(root: XmlElement, ref: string): XmlElement {
  const cell = elements(root, "c", S_NS).find((element) => element.attributes.r === ref);
  expect(cell).toBeDefined();
  return cell!;
}

function findRow(root: XmlElement, rowNumber: number): XmlElement {
  const row = elements(root, "row", S_NS).find((element) => element.attributes.r === String(rowNumber));
  expect(row).toBeDefined();
  return row!;
}

function directSpreadsheetChildren(root: XmlElement, localName: string): XmlElement[] {
  return root.children.filter((child) => child.localName === localName && child.namespaceURI === S_NS);
}

function onlyDirectSpreadsheetChild(root: XmlElement, localName: string): XmlElement {
  const children = directSpreadsheetChildren(root, localName);
  expect(children).toHaveLength(1);
  return children[0]!;
}

function relationshipId(element: XmlElement): string {
  const attributeName = Object.keys(element.attributes).find((name) => {
    const localName = name.includes(":") ? name.slice(name.indexOf(":") + 1) : name;
    return localName === "id" && element.attributeNamespaces[name] === "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  });
  expect(attributeName).toBeDefined();
  return element.attributes[attributeName!]!;
}
