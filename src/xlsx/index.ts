import { OoxmlError } from "../errors.ts";
import { OpcPackage } from "../opc/package.ts";
import {
  attribute,
  applyEdits,
  elements,
  escapeAttribute,
  escapeText,
  parseXml,
  type XmlDocument,
  type XmlElement,
} from "../xml/index.ts";

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const encoder = new TextEncoder();

type OpenInput = string | Uint8Array | ArrayBuffer;
type ScalarCellValue = string | number | boolean;
type CellEdit = { start: number; end: number; value: string };

type ParsedCell = {
  ref: string;
  element: XmlElement;
  valueElement?: XmlElement;
  formulaElement?: XmlElement;
  cell: Cell;
};

type WorksheetModel = {
  name: string;
  part: string;
  xml: string;
  document: XmlDocument;
  cells: Map<string, ParsedCell>;
};

export type Cell = BlankCell | StringCell | NumberCell | BooleanCell | FormulaCell;

export interface BaseCell {
  ref: string;
  styleId?: string;
  value?: string | number | boolean | null;
  formula?: string;
  cached?: string | number | boolean | null;
}

export interface BlankCell extends BaseCell {
  kind: "blank";
  value: null;
}

export interface StringCell extends BaseCell {
  kind: "string";
  value: string;
}

export interface NumberCell extends BaseCell {
  kind: "number";
  value: number;
}

export interface BooleanCell extends BaseCell {
  kind: "boolean";
  value: boolean;
}

export interface FormulaCell extends BaseCell {
  kind: "formula";
  formula: string;
  cached: string | number | boolean | null;
}

/**
 * Bun-native XLSX reader/editor for existing worksheets.
 *
 * Contracts:
 * - workbook/sheet discovery follows OOXML relationships, not guessed filenames;
 * - reads shared strings, inline strings, numerics, booleans and formula caches;
 * - value edits are conservative: only the touched worksheet XML and, when needed,
 *   recalculation metadata are rewritten;
 * - Bun does not calculate formulas here, so every formula cache is invalidated on
 *   dependent workbooks and Excel is asked to recalculate on open.
 */
export class Workbook {
  readonly package: OpcPackage;

  private readonly sourcePath?: string;
  private workbookPart = "";
  private workbookXml = "";
  private workbookDocument!: XmlDocument;
  private sharedStringsPart?: string;
  private sharedStrings: string[] = [];
  private readonly sheetsByName = new Map<string, WorksheetModel>();
  private sheetOrder: string[] = [];

  private constructor(pkg: OpcPackage, sourcePath?: string) {
    this.package = pkg;
    this.sourcePath = sourcePath;
    this.reload();
  }

  static async open(input: OpenInput): Promise<Workbook> {
    if (typeof input === "string") {
      return new Workbook(await OpcPackage.open(input), input);
    }

    const bytes = input instanceof Uint8Array
      ? input.slice()
      : new Uint8Array(input.slice(0));
    return new Workbook(await OpcPackage.open(bytes));
  }

  get sheetnames(): string[] {
    return [...this.sheetOrder];
  }

  /** Returns a worksheet handle backed by the current workbook state. */
  worksheet(name: string): Worksheet {
    this.requireSheet(name);
    return new Worksheet(this, name);
  }

  /** @internal Worksheet facade entrypoint. */
  readCell(sheetName: string, reference: string): Cell | undefined {
    const sheet = this.requireSheet(sheetName);
    return cloneCell(sheet.cells.get(normalizeCellReference(reference))?.cell);
  }

  /**
   * @internal Worksheet facade entrypoint.
   *
   * Edits are transactional. Unsupported formula topologies refuse before any ZIP
   * member changes, preserving byte identity for later saves.
   */
  writeCell(sheetName: string, reference: string, value: ScalarCellValue): void {
    const sheet = this.requireSheet(sheetName);
    const cell = sheet.cells.get(normalizeCellReference(reference));
    if (!cell) {
      throw new OoxmlError(
        "xlsx-cell-missing",
        `Missing cell ${reference.toUpperCase()} in worksheet ${sheetName}`,
      );
    }

    // Validate the selected cell first to retain its specific formula refusal.
    const valueEdits = buildValueEdits(cell, value);
    // Array/data-table followers may have cached values without their own <f>.
    // Clearing only anchors would silently leave those answers stale. Until the
    // range engine owns every result cell, refuse the whole value edit up front.
    for (const model of this.sheetsByName.values()) {
      for (const formula of elements(model.document, "f", S_NS)) {
        if (formula.attributes.t === "array" || formula.attributes.t === "dataTable") {
          throw new OoxmlError("xlsx-cache-topology-unsupported",
            `Cannot invalidate ${formula.attributes.t} result ranges in worksheet ${model.name}`);
        }
      }
    }

    const formulasPresent = [...this.sheetsByName.values()].some((model) =>
      [...model.cells.values()].some((entry) => entry.formulaElement !== undefined)
    );

    this.package.transaction(() => {
      for (const model of this.sheetsByName.values()) {
        const edits: CellEdit[] = [];
        if (model.name === sheetName) {
          edits.push(...valueEdits);
        }
        if (formulasPresent) {
          edits.push(...buildFormulaCacheInvalidationEdits(model));
        }
        if (edits.length > 0) {
          const nextXml = applyEdits(model.xml, edits);
          this.package.set(model.part, encoder.encode(nextXml));
        }
      }

      if (formulasPresent) {
        const nextWorkbookXml = ensureRecalculationFlags(this.workbookXml, this.workbookDocument);
        if (nextWorkbookXml !== this.workbookXml) {
          this.package.set(this.workbookPart, encoder.encode(nextWorkbookXml));
        }
      }
    });

    this.reload();
  }

  toBytes(): Uint8Array {
    return this.package.toBytes();
  }

  async save(path?: string): Promise<void> {
    const targetPath = path ?? this.sourcePath;
    if (!targetPath) {
      throw new OoxmlError(
        "xlsx-save-path-required",
        "Saving a workbook opened from bytes requires an explicit path",
      );
    }
    await this.package.save(targetPath);
  }

  private reload(): void {
    this.workbookPart = this.package.mainPart();
    this.workbookXml = this.package.text(this.workbookPart);
    this.workbookDocument = parseXml(this.workbookXml);

    const relationships = new Map(
      this.package.relationships(this.workbookPart).map((relationship) => [relationship.id, relationship]),
    );

    this.sharedStringsPart = this.package.related(this.workbookPart, "sharedStrings");
    this.sharedStrings = this.sharedStringsPart
      ? parseSharedStrings(this.package.text(this.sharedStringsPart))
      : [];

    this.sheetsByName.clear();
    this.sheetOrder = [];

    for (const sheetElement of elements(this.workbookDocument.root, "sheet", S_NS)) {
      const name = sheetElement.attributes.name;
      const relationshipId = attribute(sheetElement, "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
      if (!name || !relationshipId) {
        throw new OoxmlError("xlsx-workbook-invalid", "Workbook sheet entry is missing name or an officeDocument relationship id");
      }
      if (this.sheetsByName.has(name)) {
        throw new OoxmlError("xlsx-worksheet-duplicate", `Duplicate worksheet name ${name}`);
      }

      const relationship = relationships.get(relationshipId);
      if (!relationship?.resolved || relationship.type !== "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet") {
        throw new OoxmlError(
          "xlsx-workbook-invalid",
          `Workbook sheet ${name} references missing worksheet relationship ${relationshipId}`,
        );
      }

      const xml = this.package.text(relationship.resolved);
      const document = parseXml(xml);
      const cells = parseWorksheetCells(document, this.sharedStrings);
      this.sheetOrder.push(name);
      this.sheetsByName.set(name, {
        name,
        part: relationship.resolved,
        xml,
        document,
        cells,
      });
    }
  }

  private requireSheet(name: string): WorksheetModel {
    const sheet = this.sheetsByName.get(name);
    if (!sheet) {
      throw new OoxmlError("xlsx-worksheet-missing", `Missing worksheet ${name}`);
    }
    return sheet;
  }
}

/** Worksheet facade over an OOXML worksheet part. */
export class Worksheet {
  constructor(
    private readonly workbookRef: Workbook,
    readonly name: string,
  ) {}

  getCell(reference: string): Cell | undefined {
    return this.workbookRef.readCell(this.name, reference);
  }

  /**
   * Sets a scalar value on an existing simple cell.
   *
   * Shared and array formulas are refused before mutation because replacing those
   * XML structures safely would require a full formula topology editor.
   */
  setCellValue(reference: string, value: ScalarCellValue): void {
    this.workbookRef.writeCell(this.name, reference, value);
  }
}

function parseSharedStrings(xml: string): string[] {
  const document = parseXml(xml);
  if (document.root.localName !== "sst" || document.root.namespaceURI !== S_NS) {
    throw new OoxmlError("xlsx-shared-strings-invalid", "Invalid sharedStrings.xml root element");
  }
  return elements(document.root, "si", S_NS).map((item) => extractStringText(item));
}

function parseWorksheetCells(document: XmlDocument, sharedStrings: string[]): Map<string, ParsedCell> {
  if (document.root.localName !== "worksheet" || document.root.namespaceURI !== S_NS) {
    throw new OoxmlError("xlsx-worksheet-invalid", "Invalid worksheet root element");
  }

  const cells = new Map<string, ParsedCell>();
  for (const element of elements(document.root, "c", S_NS)) {
    const ref = normalizeCellReference(element.attributes.r ?? failCellReference());
    if (cells.has(ref)) {
      throw new OoxmlError("xlsx-cell-duplicate", `Duplicate cell reference ${ref}`);
    }
    const parsed = parseCell(element, sharedStrings, ref);
    if (parsed) {
      cells.set(ref, parsed);
    }
  }
  return cells;
}

function parseCell(element: XmlElement, sharedStrings: string[], ref: string): ParsedCell | undefined {
  const styleId = element.attributes.s;
  const type = element.attributes.t;
  const formulaElement = directChild(element, "f");
  const valueElement = directChild(element, "v");
  const inlineStringElement = directChild(element, "is");

  if (formulaElement) {
    return {
      ref,
      element,
      formulaElement,
      valueElement,
      cell: {
        kind: "formula",
        ref,
        styleId,
        formula: formulaElement.text,
        cached: parseFormulaCachedValue(type, valueElement?.text, sharedStrings),
      },
    };
  }

  if (type === "inlineStr") {
    return {
      ref,
      element,
      valueElement: inlineStringElement,
      cell: {
        kind: "string",
        ref,
        styleId,
        value: inlineStringElement ? extractStringText(inlineStringElement) : "",
      },
    };
  }

  if (type === "s") {
    const index = parseSharedStringIndex(valueElement?.text ?? "", ref);
    const value = sharedStrings[index];
    if (value === undefined) {
      throw new OoxmlError(
        "xlsx-shared-string-missing",
        `Shared string index ${index} is missing for cell ${ref}`,
      );
    }
    return {
      ref,
      element,
      valueElement,
      cell: {
        kind: "string",
        ref,
        styleId,
        value,
      },
    };
  }

  if (type === "b") {
    return {
      ref,
      element,
      valueElement,
      cell: {
        kind: "boolean",
        ref,
        styleId,
        value: parseBooleanValue(valueElement?.text ?? "", ref),
      },
    };
  }

  if (type === "str") {
    return {
      ref,
      element,
      valueElement,
      cell: {
        kind: "string",
        ref,
        styleId,
        value: valueElement?.text ?? "",
      },
    };
  }

  const rawValue = valueElement?.text;
  if (rawValue === undefined || rawValue === "") {
    return {
      ref,
      element,
      valueElement,
      cell: {
        kind: "blank",
        ref,
        styleId,
        value: null,
      },
    };
  }

  return {
    ref,
    element,
    valueElement,
    cell: {
      kind: "number",
      ref,
      styleId,
      value: parseNumericValue(rawValue, ref),
    },
  };
}

function parseFormulaCachedValue(
  type: string | undefined,
  rawValue: string | undefined,
  sharedStrings: string[],
): string | number | boolean | null {
  if (rawValue === undefined || rawValue === "") {
    return null;
  }

  if (type === "b") {
    return parseBooleanValue(rawValue, "formula cache");
  }
  if (type === "s") {
    const index = parseSharedStringIndex(rawValue, "formula cache");
    const value = sharedStrings[index];
    if (value === undefined) {
      throw new OoxmlError(
        "xlsx-shared-string-missing",
        `Shared string index ${index} is missing for a formula cache`,
      );
    }
    return value;
  }
  if (type === "str" || type === "inlineStr" || type === "e") {
    return rawValue;
  }
  return parseNumericValue(rawValue, "formula cache");
}

function buildValueEdits(cell: ParsedCell, value: ScalarCellValue): CellEdit[] {
  const formulaType = cell.formulaElement?.attributes.t;
  if (formulaType === "shared") {
    throw new OoxmlError(
      "xlsx-shared-formula-edit-unsupported",
      `Editing shared formula cell ${cell.ref} is not supported`,
    );
  }
  if (formulaType === "array") {
    throw new OoxmlError(
      "xlsx-array-formula-edit-unsupported",
      `Editing array formula cell ${cell.ref} is not supported`,
    );
  }
  if (cell.formulaElement) {
    throw new OoxmlError(
      "xlsx-formula-edit-unsupported",
      `Editing formula cell ${cell.ref} is not supported`,
    );
  }

  const attributes = { ...cell.element.attributes };
  let innerXml = "";

  if (typeof value === "string") {
    attributes.t = "inlineStr";
    innerXml = renderInlineString(cell.element, value);
  } else if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      throw new OoxmlError("xlsx-value-invalid", `Cell ${cell.ref} cannot store a non-finite number`);
    }
    attributes.t = "n";
    innerXml = renderValueElement(cell.element, formatNumber(value));
  } else {
    attributes.t = "b";
    innerXml = renderValueElement(cell.element, value ? "1" : "0");
  }

  if (cell.element.selfClosing) {
    return [{
      start: cell.element.start,
      end: cell.element.end,
      value: `<${cell.element.name}${renderAttributes(attributes)}>${innerXml}</${cell.element.name}>`,
    }];
  }

  return [
    {
      start: cell.element.start,
      end: cell.element.openEnd,
      value: `<${cell.element.name}${renderAttributes(attributes)}>`,
    },
    {
      start: cell.element.openEnd,
      end: cell.element.closeStart,
      value: innerXml,
    },
  ];
}

function buildFormulaCacheInvalidationEdits(sheet: WorksheetModel): CellEdit[] {
  const edits: CellEdit[] = [];
  for (const parsed of sheet.cells.values()) {
    if (!parsed.formulaElement || !parsed.valueElement) {
      continue;
    }
    if (parsed.valueElement.selfClosing || parsed.valueElement.text === "") {
      continue;
    }
    edits.push({
      start: parsed.valueElement.openEnd,
      end: parsed.valueElement.closeStart,
      value: "",
    });
  }
  return edits;
}

function ensureRecalculationFlags(workbookXml: string, document: XmlDocument): string {
  const calcPr = directChild(document.root, "calcPr");
  if (calcPr) {
    const attributes = {
      ...calcPr.attributes,
      calcMode: "auto",
      fullCalcOnLoad: "1",
      forceFullCalc: "1",
    };
    return applyEdits(workbookXml, [{
      start: calcPr.start,
      end: calcPr.openEnd,
      value: calcPr.selfClosing
        ? `<${calcPr.name}${renderAttributes(attributes)}/>`
        : `<${calcPr.name}${renderAttributes(attributes)}>`,
    }]);
  }

  const calcPrName = qualifiedName(document.root, "calcPr");
  return applyEdits(workbookXml, [{
    start: document.root.closeStart,
    end: document.root.closeStart,
    value: `<${calcPrName} calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>`,
  }]);
}

function directChild(element: XmlElement, localName: string): XmlElement | undefined {
  return element.children.find((child) => child.localName === localName && child.namespaceURI === S_NS);
}

function extractStringText(element: XmlElement): string {
  if (element.localName === "t" && element.namespaceURI === S_NS) {
    return element.text;
  }

  let sawSpreadsheetChild = false;
  const parts: string[] = [];
  for (const child of element.children) {
    if (child.namespaceURI !== S_NS) {
      continue;
    }
    sawSpreadsheetChild = true;
    if (child.localName === "t") {
      parts.push(child.text);
    } else if (child.localName === "r") {
      parts.push(extractStringText(child));
    }
  }

  return sawSpreadsheetChild ? parts.join("") : element.text;
}

function normalizeCellReference(reference: string): string {
  const normalized = reference.trim().toUpperCase();
  if (!/^[A-Z]+[1-9][0-9]*$/.test(normalized)) {
    throw new OoxmlError("xlsx-cell-reference-invalid", `Invalid cell reference ${reference}`);
  }
  return normalized;
}

function parseSharedStringIndex(rawValue: string, label: string): number {
  if (!/^[0-9]+$/.test(rawValue)) {
    throw new OoxmlError("xlsx-cell-invalid", `Invalid shared string index '${rawValue}' in ${label}`);
  }
  return Number(rawValue);
}

function parseBooleanValue(rawValue: string, label: string): boolean {
  if (rawValue === "1" || rawValue.toLowerCase() === "true") {
    return true;
  }
  if (rawValue === "0" || rawValue.toLowerCase() === "false") {
    return false;
  }
  throw new OoxmlError("xlsx-cell-invalid", `Invalid boolean value '${rawValue}' in ${label}`);
}

function parseNumericValue(rawValue: string, label: string): number {
  const numeric = Number(rawValue);
  if (!Number.isFinite(numeric)) {
    throw new OoxmlError("xlsx-cell-invalid", `Invalid numeric value '${rawValue}' in ${label}`);
  }
  return numeric;
}

function renderInlineString(cellElement: XmlElement, value: string): string {
  const preserve = /^\s|\s$/.test(value);
  const isName = qualifiedName(cellElement, "is");
  const textName = qualifiedName(cellElement, "t");
  return preserve
    ? `<${isName}><${textName} xml:space="preserve">${escapeText(value)}</${textName}></${isName}>`
    : `<${isName}><${textName}>${escapeText(value)}</${textName}></${isName}>`;
}

function renderValueElement(cellElement: XmlElement, value: string): string {
  const valueName = qualifiedName(cellElement, "v");
  return `<${valueName}>${escapeText(value)}</${valueName}>`;
}

function renderAttributes(attributes: Record<string, string>): string {
  return Object.entries(attributes)
    .map(([name, value]) => ` ${name}="${escapeAttribute(value)}"`)
    .join("");
}

function qualifiedName(element: XmlElement, localName: string): string {
  const separator = element.name.indexOf(":");
  return separator === -1 ? localName : `${element.name.slice(0, separator)}:${localName}`;
}

function formatNumber(value: number): string {
  return Object.is(value, -0) ? "-0" : String(value);
}

function cloneCell(cell: Cell | undefined): Cell | undefined {
  return cell ? { ...cell } : undefined;
}

function failCellReference(): never {
  throw new OoxmlError("xlsx-cell-invalid", "Worksheet cell is missing its r reference");
}
