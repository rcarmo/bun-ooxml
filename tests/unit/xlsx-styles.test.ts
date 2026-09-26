import { describe, expect, test } from "bun:test";
import { join } from "node:path";

import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { elements, parseXml, type XmlElement } from "../../src/xml/index.ts";
import { Workbook } from "../../src/xlsx/index.ts";
import { setCellWrapText } from "../../src/xlsx/styles.ts";

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const encoder = new TextEncoder();

describe("setCellWrapText", () => {
  test("creates a wrapped style closure for the shared default-style fixture and survives reopen", async () => {
    const fixturePath = join(import.meta.dir, "../../docs/contracts/shared-v2/pack/fixtures/default-style.xlsx");
    const originalBytes = new Uint8Array(await Bun.file(fixturePath).arrayBuffer());
    const originalParts = readZip(originalBytes);
    const workbook = await Workbook.open(originalBytes);

    workbook.worksheet("Sheet").setCellValue("A1", "first\nsecond");
    setCellWrapText(workbook, "Sheet", "A1", true);
    setCellWrapText(workbook, "Sheet", "A1", true);

    const bytes = workbook.toBytes();
    const parts = readZip(bytes);
    const reopened = await Workbook.open(bytes);
    const styles = parseXml(decodePart(parts, "xl/styles.xml"));
    const sheet = parseXml(decodePart(parts, "xl/worksheets/sheet1.xml"));
    const cell = findCell(sheet.root, "A1");
    const xfs = directSpreadsheetChildren(onlySpreadsheetElement(styles.root, "cellXfs"), "xf");
    const wrappedXf = xfs[1];
    const wrappedAlignment = wrappedXf ? directSpreadsheetChild(wrappedXf, "alignment") : undefined;

    expect(reopened.worksheet("Sheet").getCell("A1")).toMatchObject({
      kind: "string",
      value: "first\nsecond",
      styleId: "1",
    });
    expect(cell.attributes.s).toBe("1");
    expect(xfs).toHaveLength(2);
    expect(wrappedXf?.attributes.fontId).toBe("0");
    expect(wrappedXf?.attributes.fillId).toBe("0");
    expect(wrappedXf?.attributes.borderId).toBe("0");
    expect(wrappedXf?.attributes.xfId).toBe("0");
    expect(wrappedAlignment?.attributes.wrapText).toBe("1");
    expect(parts.get("[Content_Types].xml")).toEqual(originalParts.get("[Content_Types].xml"));
    expect(parts.get("xl/_rels/workbook.xml.rels")).toEqual(originalParts.get("xl/_rels/workbook.xml.rels"));
    expect(parts.get("xl/theme/theme1.xml")).toEqual(originalParts.get("xl/theme/theme1.xml"));
    expect(parts.get("xl/workbook.xml")).toEqual(originalParts.get("xl/workbook.xml"));
    expect(workbook.package.diff()).toEqual({
      added: [],
      changed: ["xl/styles.xml", "xl/worksheets/sheet1.xml"],
      removed: [],
    });
  });

  test("preserves namespace prefixes, reuses the appended equivalent xf and leaves formula caches intact", async () => {
    const originalBytes = createPrefixedStylesWorkbook();
    const originalParts = readZip(originalBytes);
    const workbook = await Workbook.open(originalBytes);

    setCellWrapText(workbook, "Prefixed", "A1", true);
    setCellWrapText(workbook, "Prefixed", "B1", true);
    setCellWrapText(workbook, "Prefixed", "B1", true);

    const bytes = workbook.toBytes();
    const parts = readZip(bytes);
    const stylesXml = decodePart(parts, "xl/styles.xml");
    const sheetXml = decodePart(parts, "xl/worksheets/sheet1.xml");
    const styles = parseXml(stylesXml);
    const sheet = parseXml(sheetXml);
    const xfs = directSpreadsheetChildren(onlySpreadsheetElement(styles.root, "cellXfs"), "xf");
    const appended = xfs[1];
    const appendedAlignment = appended ? directSpreadsheetChild(appended, "alignment") : undefined;

    expect(stylesXml).toContain('<x:cellXfs count="2">');
    expect(appended?.name).toBe("x:xf");
    expect(appendedAlignment?.name).toBe("x:alignment");
    expect(appendedAlignment?.attributes.wrapText).toBe("1");
    expect(xfs).toHaveLength(2);
    expect(findCell(sheet.root, "A1").attributes.s).toBe("1");
    expect(findCell(sheet.root, "B1").attributes.s).toBe("1");
    expect(sheetXml).toContain('<x:c r="C1"><x:f>A1</x:f><x:v>1</x:v></x:c>');
    expect(parts.get("xl/workbook.xml")).toEqual(originalParts.get("xl/workbook.xml"));
    expect(workbook.package.diff()).toEqual({
      added: [],
      changed: ["xl/styles.xml", "xl/worksheets/sheet1.xml"],
      removed: [],
    });
  });

  test("refuses a missing target cell", async () => {
    const workbook = await Workbook.open(createPrefixedStylesWorkbook());

    expect(() => setCellWrapText(workbook, "Prefixed", "Z9", true)).toThrow(
      expect.objectContaining({
        code: "xlsx-cell-missing",
      } satisfies Partial<OoxmlError>),
    );
    expect(workbook.package.diff()).toEqual({ added: [], changed: [], removed: [] });
  });

  test("rolls back before mutation when the cell style index is out of range", async () => {
    const originalBytes = createInvalidCellStyleIndexWorkbook();
    const workbook = await Workbook.open(originalBytes);

    expect(() => setCellWrapText(workbook, "Broken", "A1", true)).toThrow(
      expect.objectContaining({
        code: "xlsx-style-index-invalid",
      } satisfies Partial<OoxmlError>),
    );
    expect(workbook.package.diff()).toEqual({ added: [], changed: [], removed: [] });
    expect(workbook.toBytes()).toEqual(originalBytes);
  });

  test("rolls back before mutation when xf dependencies reference missing style collections", async () => {
    const originalBytes = createInvalidFontDependencyWorkbook();
    const workbook = await Workbook.open(originalBytes);

    expect(() => setCellWrapText(workbook, "Broken", "A1", true)).toThrow(
      expect.objectContaining({
        code: "xlsx-style-dependency-invalid",
      } satisfies Partial<OoxmlError>),
    );
    expect(workbook.package.diff()).toEqual({ added: [], changed: [], removed: [] });
    expect(workbook.toBytes()).toEqual(originalBytes);
  });
});

function createPrefixedStylesWorkbook(): Uint8Array {
  return createStyledWorkbook({
    sheetName: "Prefixed",
    worksheetXml: xml(`
      <x:worksheet xmlns:x="${S_NS}">
        <x:sheetData>
          <x:row r="1">
            <x:c r="A1" t="inlineStr"><x:is><x:t>first</x:t></x:is></x:c>
            <x:c r="B1" t="inlineStr"><x:is><x:t>second</x:t></x:is></x:c>
            <x:c r="C1"><x:f>A1</x:f><x:v>1</x:v></x:c>
          </x:row>
        </x:sheetData>
      </x:worksheet>
    `),
    workbookXml: xml(`
      <x:workbook xmlns:x="${S_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <x:sheets><x:sheet name="Prefixed" sheetId="1" r:id="rId1"/></x:sheets>
      </x:workbook>
    `),
    stylesXml: xml(`
      <x:styleSheet xmlns:x="${S_NS}">
        <x:fonts count="1"><x:font/></x:fonts>
        <x:fills count="1"><x:fill/></x:fills>
        <x:borders count="1"><x:border/></x:borders>
        <x:cellStyleXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></x:cellStyleXfs>
        <x:cellXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></x:cellXfs>
        <x:cellStyles count="1"><x:cellStyle name="Normal" xfId="0" builtinId="0"/></x:cellStyles>
      </x:styleSheet>
    `),
  });
}

function createInvalidCellStyleIndexWorkbook(): Uint8Array {
  return createStyledWorkbook({
    sheetName: "Broken",
    worksheetXml: xml(`
      <worksheet xmlns="${S_NS}">
        <sheetData><row r="1"><c r="A1" s="7" t="inlineStr"><is><t>broken</t></is></c></row></sheetData>
      </worksheet>
    `),
    workbookXml: xml(`
      <workbook xmlns="${S_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets><sheet name="Broken" sheetId="1" r:id="rId1"/></sheets>
      </workbook>
    `),
    stylesXml: xml(`
      <styleSheet xmlns="${S_NS}">
        <fonts count="1"><font/></fonts>
        <fills count="1"><fill/></fills>
        <borders count="1"><border/></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
      </styleSheet>
    `),
  });
}

function createInvalidFontDependencyWorkbook(): Uint8Array {
  return createStyledWorkbook({
    sheetName: "Broken",
    worksheetXml: xml(`
      <worksheet xmlns="${S_NS}">
        <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>broken</t></is></c></row></sheetData>
      </worksheet>
    `),
    workbookXml: xml(`
      <workbook xmlns="${S_NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        <sheets><sheet name="Broken" sheetId="1" r:id="rId1"/></sheets>
      </workbook>
    `),
    stylesXml: xml(`
      <styleSheet xmlns="${S_NS}">
        <fonts count="1"><font/></fonts>
        <fills count="1"><fill/></fills>
        <borders count="1"><border/></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="1"><xf numFmtId="0" fontId="9" fillId="0" borderId="0" xfId="0"/></cellXfs>
      </styleSheet>
    `),
  });
}

type StyledWorkbookParts = {
  sheetName: string;
  workbookXml: Uint8Array;
  worksheetXml: Uint8Array;
  stylesXml: Uint8Array;
};

function createStyledWorkbook(parts: StyledWorkbookParts): Uint8Array {
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
    ["xl/workbook.xml", parts.workbookXml],
    ["xl/_rels/workbook.xml.rels", xml(`
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
      </Relationships>
    `)],
    ["xl/styles.xml", parts.stylesXml],
    ["xl/worksheets/sheet1.xml", parts.worksheetXml],
  ]));
}

function decodePart(parts: Map<string, Uint8Array>, name: string): string {
  return new TextDecoder().decode(parts.get(name)!);
}

function onlySpreadsheetElement(root: XmlElement, localName: string): XmlElement {
  const matches = elements(root, localName, S_NS);
  expect(matches.length).toBe(1);
  return matches[0]!;
}

function directSpreadsheetChildren(root: XmlElement, localName: string): XmlElement[] {
  return root.children.filter((child) => child.localName === localName && child.namespaceURI === S_NS);
}

function directSpreadsheetChild(root: XmlElement, localName: string): XmlElement | undefined {
  return directSpreadsheetChildren(root, localName)[0];
}

function findCell(root: XmlElement, ref: string): XmlElement {
  const cell = elements(root, "c", S_NS).find((element) => element.attributes.r === ref);
  expect(cell).toBeDefined();
  return cell!;
}

function xml(source: string): Uint8Array {
  const body = source.replace(/^\s+|\s+$/g, "").replace(/>\s+</g, "><");
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>${body}`);
}
