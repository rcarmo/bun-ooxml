import { OoxmlError } from "../errors.ts";
import { elements, parseXml, type XmlElement } from "../xml/index.ts";
import { Slide, type TextAnchor } from "./index.ts";

const PRESENTATION_NS = "http://schemas.openxmlformats.org/presentationml/2006/main";
const DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main";

type PlaceholderKind = "title" | "subtitle";

type ParagraphAnalysis = {
  text: string;
  readable: boolean;
  replaceable: boolean;
};

type InternalSlide = {
  presentation: {
    package: {
      text(name: string): string;
    };
  };
};

export function findPlaceholderText(
  slide: Slide,
  kind: PlaceholderKind,
): { text: string; anchor: TextAnchor } {
  const xml = slideXml(slide);
  const document = parseXml(xml);
  const spTree = shapeTree(document.root);
  if (!spTree) {
    throw missingPlaceholder(slide, kind);
  }

  const matches = spTree.children.filter((child) => isMatchingDirectPlaceholder(child, kind));
  if (matches.length === 0) {
    throw missingPlaceholder(slide, kind);
  }
  if (matches.length > 1) {
    throw new OoxmlError(
      "PPTX_PLACEHOLDER_AMBIGUOUS",
      `Slide ${slide.index + 1} has ${matches.length} direct ${kind} placeholders in ${slide.partName}`,
    );
  }

  const shape = matches[0]!;
  const textBodies = shape.children.filter((child) => isElement(child, "txBody", PRESENTATION_NS));
  if (textBodies.length !== 1) {
    throw unsupportedPlaceholder(slide, kind, "must have exactly one direct text body");
  }

  const paragraphs = textBodies[0]!.children.filter((child) => isElement(child, "p", DRAWING_NS));
  if (paragraphs.length !== 1) {
    throw unsupportedPlaceholder(slide, kind, "must have exactly one direct paragraph");
  }

  const paragraph = paragraphs[0]!;
  const analysis = analyzeParagraph(paragraph);
  if (!analysis.readable) {
    throw unsupportedPlaceholder(slide, kind, "cannot be read faithfully");
  }
  if (!analysis.replaceable) {
    throw unsupportedPlaceholder(slide, kind, "cannot be edited safely");
  }
  if (analysis.text.length === 0) {
    throw unsupportedPlaceholder(slide, kind, "empty placeholder text is not supported yet");
  }

  const paragraphIndex = collectParagraphs(document.root).findIndex((entry) => entry === paragraph);
  if (paragraphIndex === -1) {
    throw unsupportedPlaceholder(slide, kind, "paragraph order cannot be anchored safely");
  }

  const inspected = slide.inspectText(`pptx.placeholders.${kind}`);
  const match = inspected[paragraphIndex];
  if (!match || match.text !== analysis.text) {
    throw unsupportedPlaceholder(slide, kind, "paragraph order drifted while acquiring the text anchor");
  }

  return {
    text: match.text,
    anchor: match.anchor,
  };
}

function slideXml(slide: Slide): string {
  const internal = slide as unknown as InternalSlide;
  return internal.presentation.package.text(slide.partName);
}

function shapeTree(root: XmlElement): XmlElement | undefined {
  if (!isElement(root, "sld", PRESENTATION_NS)) {
    return undefined;
  }
  const cSld = root.children.find((child) => isElement(child, "cSld", PRESENTATION_NS));
  return cSld?.children.find((child) => isElement(child, "spTree", PRESENTATION_NS));
}

function isMatchingDirectPlaceholder(shape: XmlElement, kind: PlaceholderKind): boolean {
  if (!isElement(shape, "sp", PRESENTATION_NS)) {
    return false;
  }
  const placeholder = directPlaceholder(shape);
  if (!placeholder) {
    return false;
  }
  const type = placeholder.attributes.type;
  return kind === "title"
    ? type === "title" || type === "ctrTitle"
    : type === "subTitle";
}

function directPlaceholder(shape: XmlElement): XmlElement | undefined {
  const nvSpPr = shape.children.find((child) => isElement(child, "nvSpPr", PRESENTATION_NS));
  const nvPr = nvSpPr?.children.find((child) => isElement(child, "nvPr", PRESENTATION_NS));
  return nvPr?.children.find((child) => isElement(child, "ph", PRESENTATION_NS));
}

function collectParagraphs(root: XmlElement): XmlElement[] {
  const paragraphs: XmlElement[] = [];
  for (const body of elements(root, "txBody")) {
    for (const child of body.children) {
      if (isElement(child, "p", DRAWING_NS)) {
        paragraphs.push(child);
      }
    }
  }
  return paragraphs;
}

function analyzeParagraph(paragraph: XmlElement): ParagraphAnalysis {
  const textParts: string[] = [];
  let readable = true;
  let replaceable = true;

  for (const child of paragraph.children) {
    if (isElement(child, "pPr", DRAWING_NS) || isElement(child, "endParaRPr", DRAWING_NS)) {
      continue;
    }

    if (isElement(child, "r", DRAWING_NS)) {
      const texts = child.children.filter((node) => isElement(node, "t", DRAWING_NS));
      const supportedChildren = child.children.every(
        (node) => isElement(node, "t", DRAWING_NS) || isElement(node, "rPr", DRAWING_NS),
      );
      if (texts.length !== 1 || !supportedChildren) {
        readable = false;
        replaceable = false;
        continue;
      }
      textParts.push(texts[0]!.text);
      continue;
    }

    if (isElement(child, "br", DRAWING_NS)) {
      const supportedChildren = child.children.every((node) => isElement(node, "rPr", DRAWING_NS));
      if (!supportedChildren) {
        readable = false;
      }
      textParts.push("\n");
      replaceable = false;
      continue;
    }

    if (isElement(child, "fld", DRAWING_NS)) {
      const texts = child.children.filter((node) => isElement(node, "t", DRAWING_NS));
      const supportedChildren = child.children.every(
        (node) => isElement(node, "t", DRAWING_NS)
          || isElement(node, "rPr", DRAWING_NS)
          || isElement(node, "endParaRPr", DRAWING_NS),
      );
      if (texts.length !== 1 || !supportedChildren) {
        readable = false;
        replaceable = false;
        continue;
      }
      textParts.push(texts[0]!.text);
      replaceable = false;
      continue;
    }

    readable = false;
    replaceable = false;
  }

  return {
    text: textParts.join(""),
    readable,
    replaceable,
  };
}

function isElement(element: XmlElement, localName: string, namespaceURI?: string): boolean {
  return element.localName === localName && (namespaceURI === undefined || element.namespaceURI === namespaceURI);
}

function missingPlaceholder(slide: Slide, kind: PlaceholderKind): OoxmlError {
  return new OoxmlError(
    "PPTX_PLACEHOLDER_NOT_FOUND",
    `Slide ${slide.index + 1} has no direct ${kind} placeholder in ${slide.partName}`,
  );
}

function unsupportedPlaceholder(slide: Slide, kind: PlaceholderKind, detail: string): OoxmlError {
  return new OoxmlError(
    "PPTX_UNSUPPORTED_TEXT_TOPOLOGY",
    `Slide ${slide.index + 1} ${kind} placeholder in ${slide.partName} ${detail}`,
  );
}
