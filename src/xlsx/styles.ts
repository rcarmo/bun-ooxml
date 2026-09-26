import { OoxmlError } from "../errors.ts";
import { attribute, applyEdits, elements, escapeAttribute, parseXml, type XmlElement } from "../xml/index.ts";
import { Workbook } from "./index.ts";

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

type StylesModel = {
  stylesPart: string;
  xml: string;
  root: XmlElement;
  cellXfsElement: XmlElement;
  cellXfs: XmlElement[];
  fonts: XmlElement[];
  fills: XmlElement[];
  borders: XmlElement[];
  cellStyleXfs: XmlElement[];
};

type CanonicalNode = {
  localName: string;
  namespaceURI: string;
  attributes: Record<string, string>;
  children: CanonicalNode[];
};

/**
 * Assigns an equivalent existing cell style or appends a cloned file xf whose
 * alignment wrapText flag matches `enabled`.
 *
 * Contracts:
 * - only existing cells are supported;
 * - omitted `s` means style index zero;
 * - the source xf and its referenced style collections must already resolve;
 * - only the worksheet part and, when needed, styles.xml are rewritten;
 * - the workbook is reloaded after a successful mutation so subsequent reads see
 *   the new style id immediately.
 */
export function setCellWrapText(
  workbook: Workbook,
  sheetName: string,
  address: string,
  enabled: boolean,
): void {
  const normalizedAddress = normalizeCellReference(address);
  const stylesPart = workbook.package.related(workbook.package.mainPart(), "styles");
  if (!stylesPart) {
    throw new OoxmlError("xlsx-styles-missing", "Workbook is missing a styles relationship");
  }

  const worksheetPart = resolveWorksheetPart(workbook, sheetName);
  let changed = false;

  workbook.package.transaction(() => {
    const stylesXml = workbook.package.text(stylesPart);
    const stylesDocument = parseXml(stylesXml);
    const styles = readStylesModel(stylesPart, stylesXml, stylesDocument.root);

    const worksheetXml = workbook.package.text(worksheetPart);
    const worksheetDocument = parseXml(worksheetXml);
    if (worksheetDocument.root.localName !== "worksheet" || worksheetDocument.root.namespaceURI !== S_NS) {
      throw new OoxmlError("xlsx-worksheet-invalid", `Invalid worksheet root element for ${sheetName}`);
    }

    const cell = elements(worksheetDocument.root, "c", S_NS)
      .find((element) => normalizeCellReference(element.attributes.r ?? failCellReference()) === normalizedAddress);
    if (!cell) {
      throw new OoxmlError(
        "xlsx-cell-missing",
        `Missing cell ${normalizedAddress} in worksheet ${sheetName}`,
      );
    }

    const currentStyleIndex = parseStyleIndex(
      cell.attributes.s ?? "0",
      `cell ${sheetName}!${normalizedAddress}`,
    );
    const currentXf = styles.cellXfs[currentStyleIndex];
    if (!currentXf) {
      throw new OoxmlError(
        "xlsx-style-index-invalid",
        `Cell ${sheetName}!${normalizedAddress} references missing cellXf ${currentStyleIndex}`,
      );
    }

    const desiredKey = canonicalKey(buildDesiredXfNode(currentXf, enabled));
    const currentKey = canonicalKey(buildExistingXfNode(currentXf));
    if (currentKey === desiredKey) {
      return;
    }

    let desiredStyleIndex = styles.cellXfs.findIndex((xf) => canonicalKey(buildExistingXfNode(xf)) === desiredKey);
    let nextStylesXml = stylesXml;

    if (desiredStyleIndex === -1) {
      desiredStyleIndex = styles.cellXfs.length;
      const clonedXfXml = renderDesiredXfXml(stylesXml, currentXf, enabled);
      nextStylesXml = appendCellXf(stylesXml, styles.cellXfsElement, clonedXfXml, desiredStyleIndex + 1);
    }

    const nextWorksheetXml = rewriteCellStyle(worksheetXml, cell, desiredStyleIndex);

    if (nextStylesXml !== stylesXml) {
      workbook.package.set(stylesPart, nextStylesXml);
      changed = true;
    }
    if (nextWorksheetXml !== worksheetXml) {
      workbook.package.set(worksheetPart, nextWorksheetXml);
      changed = true;
    }
  });

  if (changed) {
    reloadWorkbook(workbook);
  }
}

function resolveWorksheetPart(workbook: Workbook, sheetName: string): string {
  const workbookPart = workbook.package.mainPart();
  const workbookXml = parseXml(workbook.package.text(workbookPart));
  const relationships = new Map(
    workbook.package.relationships(workbookPart)
      .filter((relationship) => !relationship.external)
      .map((relationship) => [relationship.id, relationship.resolved]),
  );

  for (const sheet of elements(workbookXml.root, "sheet", S_NS)) {
    if (sheet.attributes.name !== sheetName) {
      continue;
    }
    const relationshipId = attribute(sheet, "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    if (!relationshipId) {
      throw new OoxmlError("xlsx-workbook-invalid", `Workbook sheet ${sheetName} is missing r:id`);
    }
    const resolved = relationships.get(relationshipId);
    if (!resolved) {
      throw new OoxmlError(
        "xlsx-workbook-invalid",
        `Workbook sheet ${sheetName} references missing worksheet relationship ${relationshipId}`,
      );
    }
    return resolved;
  }

  throw new OoxmlError("xlsx-worksheet-missing", `Missing worksheet ${sheetName}`);
}

function readStylesModel(stylesPart: string, xml: string, root: XmlElement): StylesModel {
  if (root.localName !== "styleSheet" || root.namespaceURI !== S_NS) {
    throw new OoxmlError("xlsx-styles-invalid", `Invalid styles root element in ${stylesPart}`);
  }

  const fontsElement = directSpreadsheetChild(root, "fonts");
  const fillsElement = directSpreadsheetChild(root, "fills");
  const bordersElement = directSpreadsheetChild(root, "borders");
  const cellStyleXfsElement = directSpreadsheetChild(root, "cellStyleXfs");
  const cellXfsElement = directSpreadsheetChild(root, "cellXfs");
  if (!cellXfsElement) {
    throw new OoxmlError("xlsx-styles-invalid", `Missing cellXfs collection in ${stylesPart}`);
  }

  const fonts = directSpreadsheetChildren(fontsElement, "font");
  const fills = directSpreadsheetChildren(fillsElement, "fill");
  const borders = directSpreadsheetChildren(bordersElement, "border");
  const cellStyleXfs = directSpreadsheetChildren(cellStyleXfsElement, "xf");
  const cellXfs = directSpreadsheetChildren(cellXfsElement, "xf");

  validateCollectionCount(fontsElement, fonts.length, `${stylesPart} fonts`);
  validateCollectionCount(fillsElement, fills.length, `${stylesPart} fills`);
  validateCollectionCount(bordersElement, borders.length, `${stylesPart} borders`);
  validateCollectionCount(cellStyleXfsElement, cellStyleXfs.length, `${stylesPart} cellStyleXfs`);
  validateCollectionCount(cellXfsElement, cellXfs.length, `${stylesPart} cellXfs`);

  if (cellXfs.length === 0) {
    throw new OoxmlError("xlsx-style-index-invalid", `Styles part ${stylesPart} has no cellXfs entries`);
  }

  const model: StylesModel = {
    stylesPart,
    xml,
    root,
    cellXfsElement,
    cellXfs,
    fonts,
    fills,
    borders,
    cellStyleXfs,
  };

  cellStyleXfs.forEach((xf, index) => validateXfDependencies(model, xf, `cellStyleXfs[${index}]`));
  cellXfs.forEach((xf, index) => validateXfDependencies(model, xf, `cellXfs[${index}]`));
  return model;
}

function validateCollectionCount(
  element: XmlElement | undefined,
  actualCount: number,
  label: string,
): void {
  const declared = element?.attributes.count;
  if (declared === undefined) {
    return;
  }
  const parsed = parseStyleIndex(declared, `${label} count`);
  if (parsed !== actualCount) {
    throw new OoxmlError(
      "xlsx-styles-invalid",
      `${label} declares ${declared} entries but contains ${actualCount}`,
    );
  }
}

function validateXfDependencies(model: StylesModel, xf: XmlElement, label: string): void {
  validateOptionalReference(xf.attributes.fontId, model.fonts.length, "fontId", label);
  validateOptionalReference(xf.attributes.fillId, model.fills.length, "fillId", label);
  validateOptionalReference(xf.attributes.borderId, model.borders.length, "borderId", label);
  validateOptionalReference(xf.attributes.xfId, model.cellStyleXfs.length, "xfId", label);
}

function validateOptionalReference(
  value: string | undefined,
  count: number,
  attributeName: string,
  label: string,
): void {
  if (value === undefined) {
    return;
  }
  const index = parseStyleIndex(value, `${label} ${attributeName}`);
  if (index >= count) {
    throw new OoxmlError(
      "xlsx-style-dependency-invalid",
      `${label} references missing ${attributeName} ${index}`,
    );
  }
}

function rewriteCellStyle(xml: string, cell: XmlElement, styleIndex: number): string {
  const attributes = { ...cell.attributes };
  if (styleIndex === 0) {
    delete attributes.s;
  } else {
    attributes.s = String(styleIndex);
  }

  if (cell.selfClosing) {
    return applyEdits(xml, [{
      start: cell.start,
      end: cell.end,
      value: `<${cell.name}${renderAttributes(attributes)}/>`,
    }]);
  }

  return applyEdits(xml, [{
    start: cell.start,
    end: cell.openEnd,
    value: `<${cell.name}${renderAttributes(attributes)}>`,
  }]);
}

function appendCellXf(xml: string, cellXfs: XmlElement, xfXml: string, nextCount: number): string {
  if (cellXfs.selfClosing) {
    throw new OoxmlError("xlsx-styles-invalid", "cellXfs cannot be self-closing when a default style is required");
  }

  const attributes = { ...cellXfs.attributes, count: String(nextCount) };
  return applyEdits(xml, [
    {
      start: cellXfs.start,
      end: cellXfs.openEnd,
      value: `<${cellXfs.name}${renderAttributes(attributes)}>`,
    },
    {
      start: cellXfs.closeStart,
      end: cellXfs.closeStart,
      value: xfXml,
    },
  ]);
}

function renderDesiredXfXml(stylesXml: string, xf: XmlElement, enabled: boolean): string {
  const alignment = directSpreadsheetChild(xf, "alignment");
  const desiredAlignmentXml = renderDesiredAlignmentXml(stylesXml, xf, alignment, enabled);
  const attributes = { ...xf.attributes };
  if (desiredAlignmentXml) {
    attributes.applyAlignment = "1";
  } else {
    delete attributes.applyAlignment;
  }

  const children: string[] = [];
  let alignmentInserted = false;
  for (const child of xf.children) {
    if (child === alignment) {
      if (desiredAlignmentXml) {
        children.push(desiredAlignmentXml);
        alignmentInserted = true;
      }
      continue;
    }
    if (!alignmentInserted && desiredAlignmentXml && shouldInsertAlignmentBefore(child)) {
      children.push(desiredAlignmentXml);
      alignmentInserted = true;
    }
    children.push(stylesXml.slice(child.start, child.end));
  }
  if (!alignmentInserted && desiredAlignmentXml) {
    children.push(desiredAlignmentXml);
  }

  if (children.length === 0) {
    return `<${xf.name}${renderAttributes(attributes)}/>`;
  }
  return `<${xf.name}${renderAttributes(attributes)}>${children.join("")}</${xf.name}>`;
}

function renderDesiredAlignmentXml(
  stylesXml: string,
  xf: XmlElement,
  alignment: XmlElement | undefined,
  enabled: boolean,
): string | undefined {
  if (!alignment && !enabled) {
    return undefined;
  }

  const attributes = { ...(alignment?.attributes ?? {}) };
  if (enabled) {
    attributes.wrapText = "1";
  } else {
    delete attributes.wrapText;
  }

  const innerXml = alignment && !alignment.selfClosing
    ? stylesXml.slice(alignment.openEnd, alignment.closeStart)
    : "";
  if (!enabled && innerXml.length === 0 && Object.keys(attributes).length === 0) {
    return undefined;
  }

  const name = alignment?.name ?? qualifiedName(xf, "alignment");
  if (innerXml.length === 0) {
    return `<${name}${renderAttributes(attributes)}/>`;
  }
  return `<${name}${renderAttributes(attributes)}>${innerXml}</${name}>`;
}

function shouldInsertAlignmentBefore(child: XmlElement): boolean {
  return child.namespaceURI === S_NS && (child.localName === "protection" || child.localName === "extLst");
}

function buildExistingXfNode(xf: XmlElement): CanonicalNode {
  const alignment = directSpreadsheetChild(xf, "alignment");
  const children = xf.children.map((child) => buildExistingNode(child));
  const attributes = normalizeXfAttributes(xf.attributes, alignment !== undefined);
  return {
    localName: xf.localName,
    namespaceURI: xf.namespaceURI,
    attributes,
    children,
  };
}

function buildExistingNode(element: XmlElement): CanonicalNode {
  const children = element.children.map((child) => buildExistingNode(child));
  const attributes = element.localName === "alignment"
    ? normalizeAlignmentAttributes(element.attributes)
    : { ...element.attributes };
  return {
    localName: element.localName,
    namespaceURI: element.namespaceURI,
    attributes,
    children,
  };
}

function buildDesiredXfNode(xf: XmlElement, enabled: boolean): CanonicalNode {
  const alignment = directSpreadsheetChild(xf, "alignment");
  const desiredAlignment = buildDesiredAlignmentNode(alignment, enabled);
  const existingChildren = xf.children.map((child) => buildExistingNode(child));
  const children: CanonicalNode[] = [];
  let alignmentInserted = false;

  for (let index = 0; index < xf.children.length; index += 1) {
    const child = xf.children[index]!;
    const built = existingChildren[index]!;
    if (child === alignment) {
      if (desiredAlignment) {
        children.push(desiredAlignment);
        alignmentInserted = true;
      }
      continue;
    }
    if (!alignmentInserted && desiredAlignment && shouldInsertAlignmentBefore(child)) {
      children.push(desiredAlignment);
      alignmentInserted = true;
    }
    children.push(built);
  }
  if (!alignmentInserted && desiredAlignment) {
    children.push(desiredAlignment);
  }

  return {
    localName: xf.localName,
    namespaceURI: xf.namespaceURI,
    attributes: normalizeXfAttributes(xf.attributes, desiredAlignment !== undefined),
    children,
  };
}

function buildDesiredAlignmentNode(alignment: XmlElement | undefined, enabled: boolean): CanonicalNode | undefined {
  if (!alignment && !enabled) {
    return undefined;
  }

  const attributes = { ...(alignment?.attributes ?? {}) };
  if (enabled) {
    attributes.wrapText = "1";
  } else {
    delete attributes.wrapText;
  }

  const normalized = normalizeAlignmentAttributes(attributes);
  const children = alignment?.children.map((child) => buildExistingNode(child)) ?? [];
  if (!enabled && Object.keys(normalized).length === 0 && children.length === 0) {
    return undefined;
  }

  return {
    localName: "alignment",
    namespaceURI: S_NS,
    attributes: normalized,
    children,
  };
}

function normalizeXfAttributes(
  attributes: Record<string, string>,
  hasAlignment: boolean,
): Record<string, string> {
  const normalized = { ...attributes };
  if (hasAlignment) {
    normalized.applyAlignment = "1";
  } else {
    delete normalized.applyAlignment;
  }
  return normalized;
}

function normalizeAlignmentAttributes(attributes: Record<string, string>): Record<string, string> {
  const normalized = { ...attributes };
  const wrapText = normalized.wrapText;
  if (wrapText === undefined) {
    return normalized;
  }
  const booleanValue = normalizeBooleanAttribute(wrapText, "alignment wrapText");
  if (booleanValue === "1") {
    normalized.wrapText = "1";
  } else {
    delete normalized.wrapText;
  }
  return normalized;
}

function canonicalKey(node: CanonicalNode): string {
  const attributes = Object.fromEntries(
    Object.entries(node.attributes).sort(([left], [right]) => left.localeCompare(right)),
  );
  return JSON.stringify({
    localName: node.localName,
    namespaceURI: node.namespaceURI,
    attributes,
    children: node.children.map((child) => canonicalKey(child)),
  });
}

function normalizeBooleanAttribute(value: string, label: string): "0" | "1" {
  const lower = value.toLowerCase();
  if (value === "1" || lower === "true") {
    return "1";
  }
  if (value === "0" || lower === "false") {
    return "0";
  }
  throw new OoxmlError("xlsx-styles-invalid", `Invalid boolean value '${value}' in ${label}`);
}

function directSpreadsheetChild(element: XmlElement | undefined, localName: string): XmlElement | undefined {
  return element?.children.find((child) => child.localName === localName && child.namespaceURI === S_NS);
}

function directSpreadsheetChildren(element: XmlElement | undefined, localName: string): XmlElement[] {
  if (!element) {
    return [];
  }
  return element.children.filter((child) => child.localName === localName && child.namespaceURI === S_NS);
}

function parseStyleIndex(rawValue: string, label: string): number {
  if (!/^[0-9]+$/.test(rawValue)) {
    throw new OoxmlError("xlsx-style-index-invalid", `Invalid style index '${rawValue}' in ${label}`);
  }
  return Number(rawValue);
}

function normalizeCellReference(reference: string): string {
  const normalized = reference.trim().toUpperCase();
  const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(normalized);
  if (!match) {
    throw new OoxmlError("xlsx-cell-reference-invalid", `Invalid cell reference ${reference}`);
  }

  let columnNumber = 0;
  for (const character of match[1]!) {
    columnNumber = columnNumber * 26 + (character.charCodeAt(0) - 64);
  }
  const rowNumber = Number(match[2]);
  if (columnNumber > 16_384 || rowNumber > 1_048_576) {
    throw new OoxmlError("xlsx-cell-reference-invalid", `Invalid cell reference ${reference}`);
  }

  return `${match[1]!}${rowNumber}`;
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

function reloadWorkbook(workbook: Workbook): void {
  (workbook as unknown as { reload: () => void }).reload();
}

function failCellReference(): never {
  throw new OoxmlError("xlsx-cell-invalid", "Worksheet cell is missing its r reference");
}
