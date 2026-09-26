import { afterEach, describe, expect, test } from "bun:test";
import { mkdtemp, mkdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";

import { runAcceptance } from "../../scripts/acceptance.ts";
import { readZip } from "../../src/opc/zip.ts";
import { OoxmlError } from "../../src/errors.ts";
import { Workbook } from "../../src/xlsx/index.ts";
import {
  bindings,
  createCrossSheetCachedFormulaWorkbook,
  createRelLinkedSharedStringsWorkbook,
} from "../acceptance/xlsx.ts";

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
    const path = join(
      import.meta.dir,
      "../../fixtures/go-ooxml/testdata/excel/multiple_sheets.xlsx",
    );
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
    const path = join(
      import.meta.dir,
      "../../fixtures/python-office-mcp-server/tests/_templates/testdata/excel/comments.xlsx",
    );
    const bytes = new Uint8Array(await Bun.file(path).arrayBuffer());
    const workbook = await Workbook.open(bytes);

    expect(workbook.sheetnames).toEqual(["Annotated"]);
    expect(workbook.worksheet("Annotated").getCell("A1")?.value).toBe("Character");
    expect(workbook.worksheet("Annotated").getCell("B1")?.value).toBe("Notes");
    expect(workbook.worksheet("Annotated").getCell("A2")?.value).toBe("Victor Frankenstein");
    expect(workbook.worksheet("Annotated").getCell("B3")?.value).toBe("The creation");
  });
});

describe("Worksheet.setCellValue", () => {
  test("preserves style attributes and unrelated parts after save and reopen", async () => {
    const path = join(import.meta.dir, "../../fixtures/go-ooxml/testdata/excel/formatting.xlsx");
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
    const workbookXml = new TextDecoder().decode(parts.get("xl/workbook.xml")!);
    const modelXml = new TextDecoder().decode(parts.get("xl/worksheets/sheet1.xml")!);
    const summaryXml = new TextDecoder().decode(parts.get("xl/worksheets/sheet2.xml")!);
    const reopened = await Workbook.open(bytes);

    expect(workbookXml).toContain('calcMode="auto"');
    expect(workbookXml).toContain('fullCalcOnLoad="1"');
    expect(workbookXml).toContain('forceFullCalc="1"');
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

  test("refuses unsupported shared formula edits without mutating the package", async () => {
    const path = join(
      import.meta.dir,
      "../../references/fixtures-ooxml/reference-assets/xlsx/tests/paper/fixtures/features/shared_formulas.xlsx",
    );
    const originalBytes = new Uint8Array(await Bun.file(path).arrayBuffer());
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
    const path = join(
      import.meta.dir,
      "../../references/fixtures-ooxml/reference-assets/xlsx/tests/paper/fixtures/features/shared_formulas.xlsx",
    );
    const originalBytes = new Uint8Array(await Bun.file(path).arrayBuffer());
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
    const featureText = await Bun.file(join(import.meta.dir, "../../features/xlsx/cells.feature")).text();
    await mkdir(join(root, "features", "xlsx"), { recursive: true });
    await Bun.write(join(root, "features/xlsx/cells.feature"), featureText);

    const report = await runAcceptance(bindings, { root });

    expect(report.status).toBe("passed");
    expect(report.inventory.features.implemented).toBe(1);
    expect(report.inventory.scenarios.implemented).toBe(5);
    expect(report.execution.cases.failed).toBe(0);
    expect(report.execution.steps.undefined).toBe(0);
  });
});

async function tempRoot(): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-xlsx-test-"));
  roots.push(root);
  return root;
}
