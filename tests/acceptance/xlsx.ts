import assert from "node:assert/strict";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { dirname, join } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { OoxmlError } from "../../src/errors.ts";
import { Workbook } from "../../src/xlsx/index.ts";

type ScenarioState = {
  sourceBytes?: Uint8Array;
  workbook?: Workbook;
  reopened?: Workbook;
  originalParts?: Map<string, Uint8Array>;
  savedBytes?: Uint8Array;
  savedParts?: Map<string, Uint8Array>;
  refusal?: unknown;
  outputPath?: string;
};

const encoder = new TextEncoder();
const projectRoot = join(import.meta.dir, "..", "..");

export function createRelLinkedSharedStringsWorkbook(): Uint8Array {
  return writeZip(
    new Map<string, Uint8Array>([
      ["[Content_Types].xml", xml(`
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
        </Types>
      `)],
      ["_rels/.rels", xml(`
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      `)],
      ["xl/workbook.xml", xml(`
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Alpha" sheetId="1" r:id="rId2"/>
            <sheet name="Beta" sheetId="2" r:id="rId1"/>
          </sheets>
          <calcPr calcId="124519"/>
        </workbook>
      `)],
      ["xl/_rels/workbook.xml.rels", xml(`
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
          <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
        </Relationships>
      `)],
      ["xl/sharedStrings.xml", xml(`
        <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
          <si><t>alpha shared</t></si>
          <si><t>beta shared</t></si>
        </sst>
      `)],
      ["xl/worksheets/sheet1.xml", xml(`
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1" t="s"><v>1</v></c>
              <c r="B1" t="n"><v>42</v></c>
              <c r="C1" t="b"><v>1</v></c>
              <c r="D1"><f>B1*2</f><v>84</v></c>
            </row>
          </sheetData>
        </worksheet>
      `)],
      ["xl/worksheets/sheet2.xml", xml(`
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1" t="s"><v>0</v></c>
              <c r="B1"><v>7</v></c>
              <c r="C1" t="b"><v>0</v></c>
              <c r="D1"><f>B1+1</f><v>8</v></c>
            </row>
          </sheetData>
        </worksheet>
      `)],
    ]),
  );
}

export function createCrossSheetCachedFormulaWorkbook(): Uint8Array {
  return writeZip(
    new Map<string, Uint8Array>([
      ["[Content_Types].xml", xml(`
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
          <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
          <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
        </Types>
      `)],
      ["_rels/.rels", xml(`
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        </Relationships>
      `)],
      ["xl/workbook.xml", xml(`
        <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <sheets>
            <sheet name="Model" sheetId="1" r:id="rId1"/>
            <sheet name="Summary" sheetId="2" r:id="rId2"/>
          </sheets>
          <calcPr calcId="124519"/>
        </workbook>
      `)],
      ["xl/_rels/workbook.xml.rels", xml(`
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
          <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
        </Relationships>
      `)],
      ["xl/worksheets/sheet1.xml", xml(`
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1"><v>1</v></c>
              <c r="B1"><f>A1+A2</f><v>3</v></c>
            </row>
            <row r="2">
              <c r="A2"><v>2</v></c>
              <c r="B2"><f>B1*2</f><v>6</v></c>
            </row>
          </sheetData>
        </worksheet>
      `)],
      ["xl/worksheets/sheet2.xml", xml(`
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData>
            <row r="1">
              <c r="A1"><f>Model!B1*3</f><v>9</v></c>
            </row>
          </sheetData>
        </worksheet>
      `)],
    ]),
  );
}

export const bindings: StepBinding[] = [
  {
    pattern: /^a synthetic two-sheet XLSX fixture with rel-linked shared strings, numbers, booleans and cached formulas$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.sourceBytes = createRelLinkedSharedStringsWorkbook();
      state.originalParts = readZip(state.sourceBytes);
    },
  },
  {
    pattern: /^the go-ooxml formatting workbook fixture$/,
    run: async (context) => {
      const state = scenarioState(context);
      const path = projectPath("fixtures/go-ooxml/testdata/excel/formatting.xlsx");
      state.sourceBytes = new Uint8Array(await Bun.file(path).arrayBuffer());
      state.originalParts = readZip(state.sourceBytes);
      state.workbook = await Workbook.open(path);
    },
  },
  {
    pattern: /^a synthetic two-sheet XLSX fixture with cross-sheet cached formulas$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.sourceBytes = createCrossSheetCachedFormulaWorkbook();
      state.originalParts = readZip(state.sourceBytes);
      state.workbook = await Workbook.open(state.sourceBytes);
    },
  },
  {
    pattern: /^the shared-formula XLSX fixture$/,
    run: async (context) => {
      const state = scenarioState(context);
      const path = projectPath("references/fixtures-ooxml/reference-assets/xlsx/tests/paper/fixtures/features/shared_formulas.xlsx");
      state.sourceBytes = new Uint8Array(await Bun.file(path).arrayBuffer());
      state.originalParts = readZip(state.sourceBytes);
      state.workbook = await Workbook.open(state.sourceBytes);
    },
  },
  {
    pattern: /^I open the workbook through the XLSX reader$/,
    run: async (context) => {
      const state = scenarioState(context);
      state.workbook ??= await Workbook.open(required(state.sourceBytes, "sourceBytes"));
    },
  },
  {
    pattern: /^I change the styled cell A2 text to "([^"]+)" and save and reopen the workbook$/,
    run: async (context, value) => {
      const state = scenarioState(context);
      const workbook = required(state.workbook, "workbook");
      workbook.worksheet("Formatted").setCellValue("A2", value);
      state.outputPath = await tempWorkbookPath("styled-edit.xlsx");
      await workbook.save(state.outputPath);
      state.savedBytes = new Uint8Array(await Bun.file(state.outputPath).arrayBuffer());
      state.savedParts = readZip(state.savedBytes);
      state.reopened = await Workbook.open(state.outputPath);
    },
  },
  {
    pattern: /^I change the input cell A1 to (\d+) and save the workbook$/,
    run: async (context, numeric) => {
      const state = scenarioState(context);
      const workbook = required(state.workbook, "workbook");
      workbook.worksheet("Model").setCellValue("A1", Number(numeric));
      state.outputPath = await tempWorkbookPath("cache-edit.xlsx");
      await workbook.save(state.outputPath);
      state.savedBytes = new Uint8Array(await Bun.file(state.outputPath).arrayBuffer());
      state.savedParts = readZip(state.savedBytes);
      state.reopened = await Workbook.open(state.savedBytes);
    },
  },
  {
    pattern: /^I try to overwrite the shared formula cell B2$/,
    run: async (context) => {
      const state = scenarioState(context);
      try {
        required(state.workbook, "workbook").worksheet("Calc").setCellValue("B2", 99);
      } catch (error) {
        state.refusal = error;
      }
      state.outputPath = await tempWorkbookPath("shared-formula-rollback.xlsx");
      await required(state.workbook, "workbook").save(state.outputPath);
      state.savedBytes = new Uint8Array(await Bun.file(state.outputPath).arrayBuffer());
      state.savedParts = readZip(state.savedBytes);
    },
  },
  {
    pattern: /^I try to overwrite the array formula cell D2$/,
    run: async (context) => {
      const state = scenarioState(context);
      try {
        required(state.workbook, "workbook").worksheet("Calc").setCellValue("D2", 99);
      } catch (error) {
        state.refusal = error;
      }
      state.outputPath = await tempWorkbookPath("array-formula-rollback.xlsx");
      await required(state.workbook, "workbook").save(state.outputPath);
      state.savedBytes = new Uint8Array(await Bun.file(state.outputPath).arrayBuffer());
      state.savedParts = readZip(state.savedBytes);
    },
  },
  {
    pattern: /^the workbook sheetnames follow workbook relationships$/,
    run: (context) => {
      const workbook = required(scenarioState(context).workbook, "workbook");
      assert.deepEqual(workbook.sheetnames, ["Alpha", "Beta"]);
    },
  },
  {
    pattern: /^cell reads include shared strings, numbers, booleans and cached formulas across worksheets$/,
    run: (context) => {
      const workbook = required(scenarioState(context).workbook, "workbook");
      const alpha = workbook.worksheet("Alpha");
      const beta = workbook.worksheet("Beta");
      assert.equal(alpha.getCell("A1")?.value, "alpha shared");
      assert.equal(alpha.getCell("B1")?.value, 7);
      assert.equal(alpha.getCell("C1")?.value, false);
      assert.equal(alpha.getCell("D1")?.formula, "B1+1");
      assert.equal(alpha.getCell("D1")?.cached, 8);
      assert.equal(beta.getCell("A1")?.value, "beta shared");
      assert.equal(beta.getCell("B1")?.value, 42);
      assert.equal(beta.getCell("C1")?.value, true);
      assert.equal(beta.getCell("D1")?.formula, "B1*2");
      assert.equal(beta.getCell("D1")?.cached, 84);
    },
  },
  {
    pattern: /^the reopened cell keeps its style attributes and new value$/,
    run: (context) => {
      const workbook = required(scenarioState(context).reopened, "reopened");
      const cell = workbook.worksheet("Formatted").getCell("A2");
      assert.equal(cell?.value, "Elizabeth Lavenza");
      assert.equal(cell?.styleId, "2");
    },
  },
  {
    pattern: /^unrelated ZIP parts still match the original bytes$/,
    run: (context) => {
      const state = scenarioState(context);
      const originalParts = required(state.originalParts, "originalParts");
      const savedParts = required(state.savedParts, "savedParts");
      for (const name of ["[Content_Types].xml", "docProps/core.xml", "xl/styles.xml", "xl/theme/theme1.xml"]) {
        equalBytes(originalParts.get(name), savedParts.get(name), name);
      }
    },
  },
  {
    pattern: /^every formula cache is removed across the worksheets$/,
    run: (context) => {
      const state = scenarioState(context);
      const savedParts = required(state.savedParts, "savedParts");
      const model = decode(savedParts.get("xl/worksheets/sheet1.xml"));
      const summary = decode(savedParts.get("xl/worksheets/sheet2.xml"));
      assert.match(model, /<f>A1\+A2<\/f>/);
      assert.match(model, /<f>B1\*2<\/f>/);
      assert.doesNotMatch(model, /<v>3<\/v>/);
      assert.doesNotMatch(model, /<v>6<\/v>/);
      assert.match(summary, /<f>Model!B1\*3<\/f>/);
      assert.doesNotMatch(summary, /<v>9<\/v>/);
      const reopened = required(state.reopened, "reopened");
      assert.equal(reopened.worksheet("Model").getCell("B1")?.cached, null);
      assert.equal(reopened.worksheet("Model").getCell("B2")?.cached, null);
      assert.equal(reopened.worksheet("Summary").getCell("A1")?.cached, null);
    },
  },
  {
    pattern: /^workbook calculation flags request full recalculation$/,
    run: (context) => {
      const state = scenarioState(context);
      const savedParts = required(state.savedParts, "savedParts");
      const workbookXml = decode(savedParts.get("xl/workbook.xml"));
      assert.match(workbookXml, /calcMode="auto"/);
      assert.match(workbookXml, /fullCalcOnLoad="1"/);
      assert.match(workbookXml, /forceFullCalc="1"/);
    },
  },
  {
    pattern: /^the shared formula edit is refused before mutation$/,
    run: (context) => {
      const refusal = scenarioState(context).refusal;
      assert.ok(refusal instanceof OoxmlError);
      assert.equal(refusal.code, "xlsx-shared-formula-edit-unsupported");
    },
  },
  {
    pattern: /^the array formula edit is refused before mutation$/,
    run: (context) => {
      const refusal = scenarioState(context).refusal;
      assert.ok(refusal instanceof OoxmlError);
      assert.equal(refusal.code, "xlsx-array-formula-edit-unsupported");
    },
  },
  {
    pattern: /^saving afterwards keeps the workbook and worksheet parts byte-identical$/,
    run: (context) => {
      const state = scenarioState(context);
      equalBytes(state.sourceBytes, state.savedBytes, "whole archive");
      const originalParts = required(state.originalParts, "originalParts");
      const savedParts = required(state.savedParts, "savedParts");
      for (const name of ["xl/workbook.xml", "xl/worksheets/sheet1.xml"]) {
        equalBytes(originalParts.get(name), savedParts.get(name), name);
      }
    },
  },
];

export default bindings;

function scenarioState(context: Record<string, unknown>): ScenarioState {
  return context.state as ScenarioState;
}

function required<T>(value: T | undefined, label: string): T {
  assert.ok(value !== undefined, `missing ${label}`);
  return value;
}

function decode(bytes: Uint8Array | undefined): string {
  return new TextDecoder().decode(required(bytes, "bytes"));
}

function equalBytes(
  left: Uint8Array | undefined,
  right: Uint8Array | undefined,
  label: string,
): void {
  const leftBytes = required(left, `${label} left`);
  const rightBytes = required(right, `${label} right`);
  assert.equal(Buffer.compare(Buffer.from(leftBytes), Buffer.from(rightBytes)), 0, `${label} bytes differ`);
}

function xml(source: string): Uint8Array {
  const body = source.replace(/^\s+|\s+$/g, "").replace(/>\s+</g, "><");
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>${body}`);
}

function projectPath(relativePath: string): string {
  return join(projectRoot, relativePath);
}

async function tempWorkbookPath(name: string): Promise<string> {
  const root = await mkdtemp(join(tmpdir(), "bun-ooxml-xlsx-"));
  return join(root, name);
}
