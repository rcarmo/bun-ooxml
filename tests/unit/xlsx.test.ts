import { afterEach, describe, expect, test } from "bun:test";
import {sharedFormulaWorkbook} from '../fixtures/native-edge-cases.ts';
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { OoxmlError } from "../../src/errors.ts";
import { readZip } from "../../src/opc/zip.ts";
import { elements, parseXml, type XmlElement } from "../../src/xml/index.ts";
import { Workbook } from "../../src/xlsx/index.ts";
import { fixturePath, fixturesRoot, F } from "../../scripts/fixture-inputs.ts";
import {
  bindings,
  createCrossSheetCachedFormulaWorkbook,
  createPhoneticRichTextWorkbook,
  createPrefixedWorkbookWithBlankCellsAndFormula,
  createRelLinkedSharedStringsWorkbook,
  createStyledBlankCellWorkbook,
} from "../acceptance/xlsx.ts";

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const roots: string[] = [];

afterEach(async () => {
  await Promise.all(roots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
});

describe("Workbook.open", () => {
  test("reads rel-linked shared strings, numerics, booleans and cached formulas across worksheets", async () => {
    const workbook = await Workbook.open(createRelLinkedSharedStringsWorkbook());

    expect(workbook.sheetnames).toEqual(["Alpha", "Beta"]);
    expect(workbook.worksheet("Alpha").getCell("A1")).toMatchObject({
      kind: "string",
      value: "alpha shared",
    });
    expect(workbook.worksheet("Alpha").getCell("B1")).toMatchObject({
      kind: "number",
      value: 7,
    });
    expect(workbook.worksheet("Alpha").getCell("C1")).toMatchObject({
      kind: "boolean",
      value: false,
    });
    expect(workbook.worksheet("Alpha").getCell("D1")).toMatchObject({
      kind: "formula",
      formula: "B1+1",
      cached: 8,
    });
    expect(workbook.worksheet("Beta").getCell("A1")).toMatchObject({
      kind: "string",
      value: "beta shared",
    });
  });

  test("reads real go-ooxml inline, numeric and boolean cells through workbook relationships", async () => {
    const path = fixturePath(F.goSheets.multipleSheets);
    const workbook = await Workbook.open(path);

    expect(workbook.sheetnames).toEqual(["Characters", "Experiments", "Summary"]);
    expect(workbook.worksheet("Characters").getCell("A2")?.value).toBe("Victor Frankenstein");
    expect(workbook.worksheet("Characters").getCell("D2")?.value).toBe(25);
    expect(workbook.worksheet("Experiments").getCell("A2")?.value).toBe("Galvanic Tests");
    expect(workbook.worksheet("Experiments").getCell("C2")?.value).toBe(47.5);
    expect(workbook.worksheet("Experiments").getCell("D2")?.value).toBe(true);
    expect(workbook.worksheet("Summary").getCell("A1")?.value).toBe("Frankenstein Research Summary");
  });

  test("reads real python-office shared strings from bytes", async () => {
    const path = fixturePath(F.officeSheets.comments);
    const bytes = new Uint8Array(await Bun.file(path).arrayBuffer());
    const workbook = await Workbook.open(bytes);

    expect(workbook.sheetnames).toEqual(["Annotated"]);
    expect(workbook.worksheet("Annotated").getCell("A1")?.value).toBe("Character");
    expect(workbook.worksheet("Annotated").getCell("B1")?.value).toBe("Notes");
    expect(workbook.worksheet("Annotated").getCell("A2")?.value).toBe("Victor Frankenstein");
    expect(workbook.worksheet("Annotated").getCell("B3")?.value).toBe("The creation");
  });

  test("excludes phonetic guides while retaining shared and inline rich text runs", async () => {
    const workbook = await Workbook.open(createPhoneticRichTextWorkbook());

    expect(workbook.sheetnames).toEqual(["Phonetics"]);
    expect(workbook.worksheet("Phonetics").getCell("A1")).toMatchObject({
      kind: "string",
      value: "Alpha Beta",
    });
    expect(workbook.worksheet("Phonetics").getCell("B1")).toMatchObject({
      kind: "string",
      value: "Gamma Delta",
    });
  });
});

describe("Worksheet.setCellValue", () => {
  test("preserves style attributes and unrelated parts after save and reopen", async () => {
    const path = fixturePath(F.goSheets.formatting);
    const originalBytes = new Uint8Array(await Bun.file(path).arrayBuffer());
    const originalParts = readZip(originalBytes);
    const workbook = await Workbook.open(path);

    workbook.worksheet("Formatted").setCellValue("A2", "Elizabeth Lavenza");

    const outDir = await tempRoot();
    const outPath = join(outDir, "styled.xlsx");
    await workbook.save(outPath);
    const savedBytes = new Uint8Array(await Bun.file(outPath).arrayBuffer());
    const savedParts = readZip(savedBytes);
    const reopened = await Workbook.open(outPath);

    expect(reopened.worksheet("Formatted").getCell("A2")).toMatchObject({
      value: "Elizabeth Lavenza",
      styleId: "2",
    });
    expect(reopened.worksheet("Formatted").getCell("D2")).toMatchObject({
      kind: "number",
      value: 25,
      styleId: "2",
    });
    expect(savedParts.get("xl/worksheets/sheet1.xml")).not.toEqual(originalParts.get("xl/worksheets/sheet1.xml"));
    expect(savedParts.get("xl/styles.xml")).toEqual(originalParts.get("xl/styles.xml"));
    expect(savedParts.get("xl/theme/theme1.xml")).toEqual(originalParts.get("xl/theme/theme1.xml"));
    expect(savedParts.get("docProps/core.xml")).toEqual(originalParts.get("docProps/core.xml"));
    expect(workbook.package.diff()).toEqual({
      added: [],
      changed: ["xl/worksheets/sheet1.xml"],
      removed: [],
    });
  });

  test("conservatively invalidates cached formulas across worksheets and sets workbook calc flags", async () => {
    const workbook = await Workbook.open(createCrossSheetCachedFormulaWorkbook());

    workbook.worksheet("Model").setCellValue("A1", 10);

    const bytes = workbook.toBytes();
    const parts = readZip(bytes);
    const workbookDocument = parseXml(decodePart(parts, "xl/workbook.xml"));
    const calcPr = onlySpreadsheetElement(workbookDocument.root, "calcPr");
    const modelXml = decodePart(parts, "xl/worksheets/sheet1.xml");
    const summaryXml = decodePart(parts, "xl/worksheets/sheet2.xml");
    const reopened = await Workbook.open(bytes);

    expect(calcPr.namespaceURI).toBe(S_NS);
    expect(calcPr.attributes.calcMode).toBe("auto");
    expect(calcPr.attributes.fullCalcOnLoad).toBe("1");
    expect(calcPr.attributes.forceFullCalc).toBe("1");
    expect(modelXml).toContain("<f>A1+A2</f>");
    expect(modelXml).toContain("<f>B1*2</f>");
    expect(modelXml).not.toContain("<v>3</v>");
    expect(modelXml).not.toContain("<v>6</v>");
    expect(summaryXml).toContain("<f>Model!B1*3</f>");
    expect(summaryXml).not.toContain("<v>9</v>");
    expect(reopened.worksheet("Model").getCell("B1")?.cached).toBeNull();
    expect(reopened.worksheet("Model").getCell("B2")?.cached).toBeNull();
    expect(reopened.worksheet("Summary").getCell("A1")?.cached).toBeNull();
    expect(workbook.package.diff()).toEqual({
      added: [],
      changed: ["xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"],
      removed: [],
    });
  });

  test("qualifies inserted spreadsheet nodes in prefixed workbook and worksheet edits without duplicating calcPr", async () => {
    const workbook = await Workbook.open(createPrefixedWorkbookWithBlankCellsAndFormula());

    workbook.worksheet("Prefixed").setCellValue("A1", "Alpha");
    workbook.worksheet("Prefixed").setCellValue("B1", 7);

    const bytes = workbook.toBytes();
    const parts = readZip(bytes);
    const workbookDocument = parseXml(decodePart(parts, "xl/workbook.xml"));
    const calcPr = onlySpreadsheetElement(workbookDocument.root, "calcPr");
    const sheetDocument = parseXml(decodePart(parts, "xl/worksheets/sheet1.xml"));
    const a1Inline = onlySpreadsheetElement(findCell(sheetDocument.root, "A1"), "is");
    const a1Text = onlySpreadsheetElement(a1Inline, "t");
    const b1Value = onlySpreadsheetElement(findCell(sheetDocument.root, "B1"), "v");
    const reopened = await Workbook.open(bytes);

    expect(calcPr.name).toBe("x:calcPr");
    expect(calcPr.namespaceURI).toBe(S_NS);
    expect(calcPr.attributes.calcMode).toBe("auto");
    expect(calcPr.attributes.fullCalcOnLoad).toBe("1");
    expect(calcPr.attributes.forceFullCalc).toBe("1");
    expect(a1Inline.name).toBe("x:is");
    expect(a1Inline.namespaceURI).toBe(S_NS);
    expect(a1Text.name).toBe("x:t");
    expect(a1Text.namespaceURI).toBe(S_NS);
    expect(a1Text.text).toBe("Alpha");
    expect(b1Value.name).toBe("x:v");
    expect(b1Value.namespaceURI).toBe(S_NS);
    expect(b1Value.text).toBe("7");
    expect(reopened.worksheet("Prefixed").getCell("A1")).toMatchObject({
      value: "Alpha",
      styleId: "1",
    });
    expect(reopened.worksheet("Prefixed").getCell("B1")).toMatchObject({
      value: 7,
      styleId: "2",
    });
  });

  test("reads styled blank cells as null and preserves their style when edited", async () => {
    const workbook = await Workbook.open(createStyledBlankCellWorkbook());

    expect(workbook.worksheet("StyledBlank").getCell("A1")).toMatchObject({
      kind: "blank",
      value: null,
      styleId: "1",
    });

    workbook.worksheet("StyledBlank").setCellValue("A1", "filled");

    const reopened = await Workbook.open(workbook.toBytes());

    expect(reopened.worksheet("StyledBlank").getCell("A1")).toMatchObject({
      kind: "string",
      value: "filled",
      styleId: "1",
    });
  });

  test("refuses unsupported shared formula edits without mutating the package", async () => {
    const originalBytes = await sharedFormulaWorkbook();
    const workbook = await Workbook.open(originalBytes);

    expect(() => workbook.worksheet("Calc").setCellValue("B2", 99)).toThrow(
      expect.objectContaining({
        code: "xlsx-shared-formula-edit-unsupported",
      } satisfies Partial<OoxmlError>),
    );
    expect(workbook.package.diff()).toEqual({ added: [], changed: [], removed: [] });
    expect(workbook.toBytes()).toEqual(originalBytes);
  });

  test("refuses unsupported array formula edits without mutating the package", async () => {
    const originalBytes = await sharedFormulaWorkbook();
    const workbook = await Workbook.open(originalBytes);

    expect(() => workbook.worksheet("Calc").setCellValue("D2", 99)).toThrow(
      expect.objectContaining({
        code: "xlsx-array-formula-edit-unsupported",
      } satisfies Partial<OoxmlError>),
    );
    expect(workbook.package.diff()).toEqual({ added: [], changed: [], removed: [] });
    expect(workbook.toBytes()).toEqual(originalBytes);
  });
});

describe("xlsx acceptance feature", () => {
  test("executes the implemented xlsx scenarios with exported bindings", async () => {
    const root = await tempRoot();
    for(const path of ["workflows/xlsx/cells.feature","workflows/xlsx/formula-cache.feature"])await Bun.write(join(root,'references/fixtures-ooxml',path),await Bun.file(join(fixturesRoot(),path)).text());
    await Bun.write(join(root,'features/shared.json'),JSON.stringify({"schemaVersion":2,"features":[{"path":"references/fixtures-ooxml/workflows/xlsx/cells.feature","lifecycle":"implemented","runner":"bun","scenarioIds":["@id-xlsx-read-rel-linked-shared-strings","@id-xlsx-preserve-styled-cell-edit","@id-xlsx-prefixed-namespace-safe-edits","@id-xlsx-phonetic-guides-excluded","@id-xlsx-styled-blank-cell-editable","@id-xlsx-refuse-shared-formula-overwrite","@id-xlsx-refuse-array-formula-overwrite"]},{"path":"references/fixtures-ooxml/workflows/xlsx/formula-cache.feature","lifecycle":"implemented","runner":"bun","scenarioIds":["@id-xlsx-clear-cross-sheet-caches"]}]}));

    const report = await runAcceptance(bindings, { root });

    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(2);
    expect(report.inventory.scenarios.implemented).toBe(8);
    expect(report.execution.cases.failed).toBe(0);
    expect(report.execution.steps.undefined).toBe(0);
  });
});

async function tempRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-xlsx-test-"));
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
  const match = elements(root, "c", S_NS).find((element) => element.attributes.r === ref);
  expect(match).toBeDefined();
  return match!;
}
