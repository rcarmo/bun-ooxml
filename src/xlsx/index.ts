import { posix } from "node:path";

import { OoxmlError } from "../errors.ts";
import { OpcPackage, relationshipPath, sameBytes } from "../opc/package.ts";
import {selectCellStyle} from './cell-style.ts';
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
const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OPC_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const WORKBOOK_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const WORKSHEET_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
const STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
const WORKBOOK_REL_TYPE = `${OFFICE_REL_NS}/officeDocument`;
const WORKSHEET_REL_TYPE = `${OFFICE_REL_NS}/worksheet`;
const STYLES_REL_TYPE = `${OFFICE_REL_NS}/styles`;
const MAX_WORKSHEET_NAME_LENGTH = 31;
const MAX_COLUMN_NUMBER = 16_384;
const MAX_ROW_NUMBER = 1_048_576;
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

type ParsedReference = {
  ref: string;
  columnName: string;
  columnNumber: number;
  rowNumber: number;
};

type WorksheetStructure = {
  dimension?: XmlElement;
  sheetData: XmlElement;
  rows: WorksheetRowStructure[];
};

type WorksheetRowStructure = {
  element: XmlElement;
  rowNumber: number;
  cells: WorksheetCellStructure[];
};

type WorksheetCellStructure = {
  element: XmlElement;
  reference: ParsedReference;
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
  private styleSourceSnapshot = new Map<string,Uint8Array>();

  private constructor(pkg: OpcPackage, sourcePath?: string) {
    this.package = pkg;
    this.sourcePath = sourcePath;
    this.reload();
  }

  static create(): Workbook {
    return new Workbook(OpcPackage.fromParts(createWorkbookParts()));
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

  addWorksheet(name: string): Worksheet {
    validateWorksheetName(name, this.sheetOrder);

    const sheetsElement = directChild(this.workbookDocument.root, "sheets");
    if (!sheetsElement) {
      throw new OoxmlError("xlsx-workbook-invalid", "Workbook is missing a sheets collection");
    }

    const workbookRelationshipsPart = relationshipPath(this.workbookPart);
    const workbookRelationshipsXml = this.package.text(workbookRelationshipsPart);
    const workbookRelationshipsDocument = parseXml(workbookRelationshipsXml);
    const contentTypesXml = this.package.text("[Content_Types].xml");
    const contentTypesDocument = parseXml(contentTypesXml);

    const nextSheetId = String(nextWorksheetSheetId(this.workbookDocument));
    const nextRelationshipId = nextWorksheetRelationshipId(this.package.relationships(this.workbookPart).map((relationship) => relationship.id));
    const worksheetPart = nextWorksheetPartName(this.package.names(), this.workbookPart);
    const worksheetTarget = posix.relative(posix.dirname(this.workbookPart), worksheetPart);
    const relationshipAttribute = existingRelationshipAttributeName(elements(this.workbookDocument.root, "sheet", S_NS)[0]);
    const relationshipPrefix=relationshipAttribute.split(':')[0]!;
    const nextSheetXml = `<${qualifiedName(sheetsElement, "sheet")} xmlns:${relationshipPrefix}="${OFFICE_REL_NS}" name="${escapeAttribute(name)}" sheetId="${escapeAttribute(nextSheetId)}" ${relationshipAttribute}="${escapeAttribute(nextRelationshipId)}"/>`;
    const nextWorkbookXml = sheetsElement.selfClosing
      ? applyEdits(this.workbookXml, [{
          start: sheetsElement.start,
          end: sheetsElement.end,
          value: `<${sheetsElement.name}${renderAttributes(sheetsElement.attributes)}>${nextSheetXml}</${sheetsElement.name}>`,
        }])
      : applyEdits(this.workbookXml, [{
          start: sheetsElement.closeStart,
          end: sheetsElement.closeStart,
          value: nextSheetXml,
        }]);
    const nextWorkbookRelationshipsXml = applyEdits(workbookRelationshipsXml, [{
      start: workbookRelationshipsDocument.root.closeStart,
      end: workbookRelationshipsDocument.root.closeStart,
      value: `<${qualifiedName(workbookRelationshipsDocument.root, "Relationship")} Id="${escapeAttribute(nextRelationshipId)}" Type="${escapeAttribute(WORKSHEET_REL_TYPE)}" Target="${escapeAttribute(worksheetTarget)}"/>`,
    }]);
    const nextContentTypesXml = applyEdits(contentTypesXml, [{
      start: contentTypesDocument.root.closeStart,
      end: contentTypesDocument.root.closeStart,
      value: `<${qualifiedName(contentTypesDocument.root, "Override")} PartName="/${escapeAttribute(worksheetPart)}" ContentType="${escapeAttribute(WORKSHEET_CONTENT_TYPE)}"/>`,
    }]);

    this.package.transaction(() => {
      this.package.set("[Content_Types].xml", encoder.encode(nextContentTypesXml));
      this.package.set(workbookRelationshipsPart, encoder.encode(nextWorkbookRelationshipsXml));
      this.package.set(this.workbookPart, encoder.encode(nextWorkbookXml));
      this.package.set(worksheetPart, createWorksheetPart());
      this.package.toBytes();
    });

    this.reload();
    return this.worksheet(name);
  }

  /** @internal Existing-cell direct style assignment. */
  setCellStyle(sheetName:string, reference:string, styleIndex:number|null): {changed:number} {
    for(const [part,bytes]of this.styleSourceSnapshot){const live=this.package.get(part);if(!live||!sameBytes(live,bytes))throw new OoxmlError('xlsx-stale-workbook','Workbook model changed outside its wrapper');}
    if(typeof reference!=='string')throw new OoxmlError('xlsx-cell-invalid','Cell reference must be a string');
    const sheet=this.requireSheet(sheetName),cell=sheet.cells.get(normalizeCellReference(reference));
    if([...this.sheetsByName.values()].filter(s=>s.part===sheet.part).length!==1)throw new OoxmlError('xlsx-style-ambiguous','Several worksheets reference the selected part');
    if(!cell)throw new OoxmlError('xlsx-cell-missing','Style assignment requires an existing cell');
    const next=selectCellStyle(this.package,this.workbookPart,sheet.xml,cell.element,styleIndex);
    if(next===sheet.xml)return {changed:0};
    const document=parseXml(next),cells=parseWorksheetCells(document,this.sharedStrings);
    this.package.transaction(()=>{this.package.set(sheet.part,next);this.package.toBytes();});
    this.sheetsByName.set(sheetName,{...sheet,xml:next,document,cells});
    this.styleSourceSnapshot.set(sheet.part,this.package.get(sheet.part)!);
    return {changed:1};
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
    const normalizedReference = normalizeCellReference(reference);
    const sheet = this.requireSheet(sheetName);
    const cell = sheet.cells.get(normalizedReference);

    // Validate the selected cell first to retain its specific formula refusal.
    const valueEdits = cell
      ? buildValueEdits(cell, value)
      : buildMissingCellEdits(sheet, normalizedReference, value);
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
      const relationshipId = attribute(sheetElement, "id", OFFICE_REL_NS);
      if (!name || !relationshipId) {
        throw new OoxmlError("xlsx-workbook-invalid", "Workbook sheet entry is missing name or an officeDocument relationship id");
      }
      if (this.sheetsByName.has(name)) {
        throw new OoxmlError("xlsx-worksheet-duplicate", `Duplicate worksheet name ${name}`);
      }

      const relationship = relationships.get(relationshipId);
      if (!relationship?.resolved || relationship.type !== WORKSHEET_REL_TYPE) {
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
    const names=['_rels/.rels',this.workbookPart,relationshipPath(this.workbookPart),...[...this.sheetsByName.values()].map(s=>s.part),...(this.sharedStringsPart?[this.sharedStringsPart]:[])];
    this.styleSourceSnapshot=new Map(names.map(n=>[n,this.package.get(n)!]));
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

  /** Assign an existing cellXf index or remove the direct override. */
  setCellStyle(reference:string, styleIndex:number|null): {changed:number} {
    return this.workbookRef.setCellStyle(this.name,reference,styleIndex);
  }

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


function buildMissingCellEdits(
  sheet: WorksheetModel,
  reference: string,
  value: ScalarCellValue,
): CellEdit[] {
  const structure = inspectWorksheetStructure(sheet);
  const parsedReference = parseCellReference(reference);
  const edits = buildDimensionEdits(structure, dimensionRef(sheet, reference));
  const row = structure.rows.find((candidate) => candidate.rowNumber === parsedReference.rowNumber);
  const cellXml = renderMissingCellXml(row?.element ?? structure.sheetData, reference, value);

  if (row) {
    edits.push(...buildMissingRowCellEdits(row, parsedReference.columnNumber, cellXml));
  } else {
    edits.push(...buildMissingRowEdits(structure.sheetData, structure.rows, parsedReference.rowNumber, cellXml));
  }

  return edits;
}

function inspectWorksheetStructure(sheet: WorksheetModel): WorksheetStructure {
  const sheetData = directChild(sheet.document.root, "sheetData");
  if (!sheetData) {
    throw unsupportedWorksheetStructure(sheet.name, "sheetData is missing");
  }

  const directDimensions = sheet.document.root.children.filter((child) =>
    child.localName === "dimension" && child.namespaceURI === S_NS
  );
  if (directDimensions.length > 1) {
    throw unsupportedWorksheetStructure(sheet.name, "multiple dimension elements are not supported");
  }

  const spreadsheetRows = sheetData.children.filter((child) => child.namespaceURI === S_NS);
  if (spreadsheetRows.some((child) => child.localName !== "row")) {
    throw unsupportedWorksheetStructure(sheet.name, "sheetData contains unsupported spreadsheet children");
  }

  const rows: WorksheetRowStructure[] = [];
  let previousRowNumber = 0;

  for (const rowElement of spreadsheetRows) {
    const rowNumber = parseWorksheetRowNumber(rowElement.attributes.r, sheet.name);
    if (rowNumber <= previousRowNumber) {
      throw unsupportedWorksheetStructure(sheet.name, "rows must be stored in ascending numeric order");
    }
    previousRowNumber = rowNumber;

    const spreadsheetCells = rowElement.children.filter((child) => child.namespaceURI === S_NS);
    if (spreadsheetCells.some((child) => child.localName !== "c")) {
      throw unsupportedWorksheetStructure(sheet.name, `row ${rowNumber} contains unsupported spreadsheet children`);
    }

    const cells: WorksheetCellStructure[] = [];
    let previousColumnNumber = 0;
    for (const cellElement of spreadsheetCells) {
      const parsedReference = parseCellReference(cellElement.attributes.r ?? failCellReference());
      if (parsedReference.rowNumber !== rowNumber) {
        throw unsupportedWorksheetStructure(sheet.name, `cell ${parsedReference.ref} is in row ${rowNumber}`);
      }
      if (parsedReference.columnNumber <= previousColumnNumber) {
        throw unsupportedWorksheetStructure(sheet.name, `row ${rowNumber} cells must be stored in ascending order`);
      }
      previousColumnNumber = parsedReference.columnNumber;
      cells.push({ element: cellElement, reference: parsedReference });
    }

    rows.push({
      element: rowElement,
      rowNumber,
      cells,
    });
  }

  return {
    dimension: directDimensions[0],
    sheetData,
    rows,
  };
}

function buildDimensionEdits(structure: WorksheetStructure, ref: string): CellEdit[] {
  if (structure.dimension) {
    const attributes = {
      ...structure.dimension.attributes,
      ref,
    };
    return [{
      start: structure.dimension.start,
      end: structure.dimension.openEnd,
      value: structure.dimension.selfClosing
        ? `<${structure.dimension.name}${renderAttributes(attributes)}/>`
        : `<${structure.dimension.name}${renderAttributes(attributes)}>`,
    }];
  }

  return [{
    start: dimensionInsertionPoint(structure.sheetData.root),
    end: dimensionInsertionPoint(structure.sheetData.root),
    value: `<${qualifiedName(structure.sheetData.root, "dimension")} ref="${escapeAttribute(ref)}"/>`,
  }];
}

function dimensionInsertionPoint(root: XmlElement):number {
  // CT_Worksheet order: optional sheetPr, dimension, sheetViews, sheetFormatPr,
  // cols, sheetData... Inserting immediately before sheetData is too late.
  return root.children.find(child=>!(child.namespaceURI===S_NS&&child.localName==='sheetPr'))?.start??root.closeStart;
}

function buildMissingRowEdits(
  sheetData: XmlElement,
  rows: WorksheetRowStructure[],
  rowNumber: number,
  cellXml: string,
): CellEdit[] {
  const rowXml = `<${qualifiedName(sheetData, "row")} r="${escapeAttribute(String(rowNumber))}">${cellXml}</${qualifiedName(sheetData, "row")}>`;
  if (sheetData.selfClosing) {
    return [{
      start: sheetData.start,
      end: sheetData.end,
      value: `<${sheetData.name}${renderAttributes(sheetData.attributes)}>${rowXml}</${sheetData.name}>`,
    }];
  }

  const nextRow = rows.find((candidate) => candidate.rowNumber > rowNumber);
  const insertionPoint = nextRow ? nextRow.element.start : sheetData.closeStart;
  return [{ start: insertionPoint, end: insertionPoint, value: rowXml }];
}

function buildMissingRowCellEdits(
  row: WorksheetRowStructure,
  columnNumber: number,
  cellXml: string,
): CellEdit[] {
  if (row.element.selfClosing) {
    return [{
      start: row.element.start,
      end: row.element.end,
      value: `<${row.element.name}${renderAttributes(row.element.attributes)}>${cellXml}</${row.element.name}>`,
    }];
  }

  const nextCell = row.cells.find((candidate) => candidate.reference.columnNumber > columnNumber);
  const insertionPoint = nextCell ? nextCell.element.start : row.element.closeStart;
  return [{ start: insertionPoint, end: insertionPoint, value: cellXml }];
}

function renderMissingCellXml(source: XmlElement, reference: string, value: ScalarCellValue): string {
  const attributes: Record<string, string> = { r: reference };
  let innerXml = "";

  if (typeof value === "string") {
    attributes.t = "inlineStr";
    innerXml = renderInlineString(source, value);
  } else if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      throw new OoxmlError("xlsx-value-invalid", `Cell ${reference} cannot store a non-finite number`);
    }
    attributes.t = "n";
    innerXml = renderValueElement(source, formatNumber(value));
  } else {
    attributes.t = "b";
    innerXml = renderValueElement(source, value ? "1" : "0");
  }

  return `<${qualifiedName(source, "c")}${renderAttributes(attributes)}>${innerXml}</${qualifiedName(source, "c")}>`;
}

function dimensionRef(sheet: WorksheetModel, reference: string): string {
  const references = [...sheet.cells.keys(), reference].map((value) => parseCellReference(value));
  let minColumn=Infinity,maxColumn=0,minRow=Infinity,maxRow=0;
  for(const value of references){minColumn=Math.min(minColumn,value.columnNumber);maxColumn=Math.max(maxColumn,value.columnNumber);minRow=Math.min(minRow,value.rowNumber);maxRow=Math.max(maxRow,value.rowNumber);}
  return `${columnName(minColumn)}${minRow}:${columnName(maxColumn)}${maxRow}`;
}

function unsupportedWorksheetStructure(sheetName: string, detail: string): OoxmlError {
  return new OoxmlError(
    "xlsx-worksheet-structure-unsupported",
    `Worksheet ${sheetName} cannot be extended safely: ${detail}`,
  );
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
  return parseCellReference(reference).ref;
}

function parseCellReference(reference: string): ParsedReference {
  const normalized = reference.trim().toUpperCase();
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(normalized);
  if (!match) {
    throw new OoxmlError("xlsx-cell-reference-invalid", `Invalid cell reference ${reference}`);
  }

  const columnNameValue = match[1]!;
  const rowNumber = Number(match[2]);
  const columnNumber = columnNumberFromName(columnNameValue);
  if (rowNumber > MAX_ROW_NUMBER || columnNumber > MAX_COLUMN_NUMBER) {
    throw new OoxmlError("xlsx-cell-reference-invalid", `Invalid cell reference ${reference}`);
  }

  return {
    ref: `${columnNameValue}${rowNumber}`,
    columnName: columnNameValue,
    columnNumber,
    rowNumber,
  };
}

function parseWorksheetRowNumber(rawValue: string | undefined, sheetName: string): number {
  if (!rawValue || !/^[1-9][0-9]*$/.test(rawValue)) {
    throw unsupportedWorksheetStructure(sheetName, "row numbering is invalid");
  }

  const rowNumber = Number(rawValue);
  if (rowNumber > MAX_ROW_NUMBER) {
    throw unsupportedWorksheetStructure(sheetName, `row ${rawValue} exceeds worksheet bounds`);
  }
  return rowNumber;
}

function columnNumberFromName(name: string): number {
  let value = 0;
  for (const character of name) {
    value = value * 26 + (character.charCodeAt(0) - 64);
  }
  return value;
}

function columnName(columnNumber: number): string {
  let current = columnNumber;
  let value = "";
  while (current > 0) {
    const offset = (current - 1) % 26;
    value = String.fromCharCode(65 + offset) + value;
    current = Math.floor((current - 1) / 26);
  }
  return value;
}

function validateWorksheetName(name: string, existingNames: string[]): void {
  if (
    name.length === 0 ||
    name.length > MAX_WORKSHEET_NAME_LENGTH ||
    /^'|'$/.test(name) ||
    /[:\\/?*\[\]]/.test(name)
  ) {
    throw new OoxmlError("xlsx-worksheet-name-invalid", `Invalid worksheet name ${name}`);
  }

  const duplicate = existingNames.find((existingName) => existingName.toLowerCase() === name.toLowerCase());
  if (duplicate) {
    throw new OoxmlError("xlsx-worksheet-duplicate", `Duplicate worksheet name ${name}`);
  }
}

function nextWorksheetSheetId(document: XmlDocument): number {
  let maxSheetId = 0;
  for (const sheet of elements(document.root, "sheet", S_NS)) {
    const rawSheetId = sheet.attributes.sheetId;
    if (!rawSheetId || !/^[1-9][0-9]*$/.test(rawSheetId)) {
      throw new OoxmlError("xlsx-workbook-invalid", "Workbook sheet entry is missing a valid sheetId");
    }
    maxSheetId = Math.max(maxSheetId, Number(rawSheetId));
  }
  return maxSheetId + 1;
}

function nextWorksheetRelationshipId(existingIds: string[]): string {
  const ids = new Set(existingIds);
  for (let index = 1; ; index += 1) {
    const candidate = `rId${index}`;
    if (!ids.has(candidate)) {
      return candidate;
    }
  }
}

function nextWorksheetPartName(existingNames: string[], workbookPart: string): string {
  const existing = new Set(existingNames);
  const workbookDirectory = posix.dirname(workbookPart);
  const prefix = workbookDirectory === "." ? "" : `${workbookDirectory}/`;

  for (let index = 1; ; index += 1) {
    const candidate = `${prefix}worksheets/sheet${index}.xml`;
    if (!existing.has(candidate)) {
      return candidate;
    }
  }
}

function existingRelationshipAttributeName(sheetElement?: XmlElement): string {
  if (!sheetElement) {
    return "r:id";
  }

  for (const attributeName of Object.keys(sheetElement.attributes)) {
    const localName = attributeName.includes(":")
      ? attributeName.slice(attributeName.indexOf(":") + 1)
      : attributeName;
    if (localName === "id" && sheetElement.attributeNamespaces[attributeName] === OFFICE_REL_NS) {
      return attributeName;
    }
  }

  return "r:id";
}

function createWorkbookParts(): Map<string, Uint8Array> {
  return new Map<string, Uint8Array>([
    ["[Content_Types].xml", xml(`
      <Types xmlns="${CT_NS}">
        <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        <Default Extension="xml" ContentType="application/xml"/>
        <Override PartName="/xl/workbook.xml" ContentType="${WORKBOOK_CONTENT_TYPE}"/>
        <Override PartName="/xl/worksheets/sheet1.xml" ContentType="${WORKSHEET_CONTENT_TYPE}"/>
        <Override PartName="/xl/styles.xml" ContentType="${STYLES_CONTENT_TYPE}"/>
      </Types>
    `)],
    ["_rels/.rels", xml(`
      <Relationships xmlns="${OPC_REL_NS}">
        <Relationship Id="rId1" Type="${WORKBOOK_REL_TYPE}" Target="xl/workbook.xml"/>
      </Relationships>
    `)],
    ["xl/workbook.xml", xml(`
      <workbook xmlns="${S_NS}" xmlns:r="${OFFICE_REL_NS}">
        <sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
      </workbook>
    `)],
    ["xl/_rels/workbook.xml.rels", xml(`
      <Relationships xmlns="${OPC_REL_NS}">
        <Relationship Id="rId1" Type="${WORKSHEET_REL_TYPE}" Target="worksheets/sheet1.xml"/>
        <Relationship Id="rId2" Type="${STYLES_REL_TYPE}" Target="styles.xml"/>
      </Relationships>
    `)],
    ["xl/styles.xml", xml(`
      <styleSheet xmlns="${S_NS}">
        <fonts count="1"><font/></fonts>
        <fills count="1"><fill/></fills>
        <borders count="1"><border/></borders>
        <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
        <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
      </styleSheet>
    `)],
    ["xl/worksheets/sheet1.xml", createWorksheetPart()],
  ]);
}

function createWorksheetPart(): Uint8Array {
  return xml(`
    <worksheet xmlns="${S_NS}">
      <sheetData></sheetData>
    </worksheet>
  `);
}

function xml(source: string): Uint8Array {
  const body = source.replace(/^\s+|\s+$/g, "").replace(/>\s+</g, "><");
  return encoder.encode(`<?xml version="1.0" encoding="UTF-8"?>${body}`);
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

function renderInlineString(source: XmlElement, value: string): string {
  const preserve = /^\s|\s$/.test(value);
  const isName = qualifiedName(source, "is");
  const textName = qualifiedName(source, "t");
  return preserve
    ? `<${isName}><${textName} xml:space="preserve">${escapeText(value)}</${textName}></${isName}>`
    : `<${isName}><${textName}>${escapeText(value)}</${textName}></${isName}>`;
}

function renderValueElement(source: XmlElement, value: string): string {
  const valueName = qualifiedName(source, "v");
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
