import { describe, expect, test } from "bun:test";
import { readFileSync } from "node:fs";
import { join } from "node:path";

import { OoxmlError } from "../../src/errors.ts";
import { readZip, writeZip } from "../../src/opc/zip.ts";
import { Presentation } from "../../src/pptx/index.ts";
import { findPlaceholderText } from "../../src/pptx/placeholders.ts";

const decoder = new TextDecoder();
const encoder = new TextEncoder();
const SHARED_FIXTURE = join(
  import.meta.dir,
  "../../docs/contracts/shared-v2/pack/fixtures/title-and-subtitle.pptx",
);
const SLIDE_PART = "ppt/slides/slide1.xml";

const TITLE_PARAGRAPH = '<a:p><a:r><a:t>Original title</a:t></a:r></a:p>';
const SUBTITLE_PARAGRAPH = '<a:p><a:r><a:t>Original subtitle</a:t></a:r></a:p>';
const TITLE_SHAPE = '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="ctrTitle" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr /><a:lstStyle />' + TITLE_PARAGRAPH + '</p:txBody></p:sp>';
const SUBTITLE_SHAPE = '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Subtitle 2" /><p:cNvSpPr><a:spLocks noGrp="1" /></p:cNvSpPr><p:nvPr><p:ph type="subTitle" idx="1" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr /><a:lstStyle />' + SUBTITLE_PARAGRAPH + '</p:txBody></p:sp>';
const BREAK_FIELD_TITLE_PARAGRAPH = '<a:p><a:r><a:rPr b="1" /><a:t>Chapter</a:t></a:r><a:br><a:rPr lang="en-US" /></a:br><a:fld id="{00000000-0000-0000-0000-000000000007}" type="slidenum"><a:rPr i="1" /><a:t>7</a:t></a:fld><a:r><a:rPr u="sng" /><a:t> Notes</a:t></a:r></a:p>';
const PLAIN_TEXT_SHAPE = '<p:sp><p:nvSpPr><p:cNvPr id="41" name="Plain Text" /><p:cNvSpPr /><p:nvPr /></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr /><a:lstStyle /><a:p><a:r><a:t>Lead in</a:t></a:r></a:p></p:txBody></p:sp>';
const NESTED_TITLE_GROUP = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="51" name="Nested Group" /><p:cNvGrpSpPr /><p:nvPr /></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /><a:chOff x="0" y="0" /><a:chExt cx="0" cy="0" /></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="52" name="Nested Title" /><p:cNvSpPr /><p:nvPr><p:ph type="title" /></p:nvPr></p:nvSpPr><p:spPr /><p:txBody><a:bodyPr /><a:lstStyle /><a:p><a:r><a:t>Nested title</a:t></a:r></a:p></p:txBody></p:sp></p:grpSp>';

describe("pptx placeholder resolution", () => {
  test("finds shared title and subtitle without mutating the package", async () => {
    const presentation = await Presentation.open(await sharedFixtureBytes());
    const slide = required(presentation.slides[0], "slide 1");

    const title = findPlaceholderText(slide, "title");
    const subtitle = findPlaceholderText(slide, "subtitle");

    expect(title.text).toBe("Original title");
    expect(subtitle.text).toBe("Original subtitle");
    expect(title.anchor.paragraphIndex).toBe(0);
    expect(subtitle.anchor.paragraphIndex).toBe(1);
    expect(presentation.package.diff()).toEqual({ added: [], changed: [], removed: [] });

    slide.replaceTextAt(title.anchor, title.text, "Changed title");
    const refreshedSubtitle = findPlaceholderText(slide, "subtitle");
    slide.replaceTextAt(refreshedSubtitle.anchor, refreshedSubtitle.text, "Changed subtitle");
    expect(slide.inspectText("unit.placeholders.shared").map((paragraph) => paragraph.text)).toEqual([
      "Changed title",
      "Changed subtitle",
    ]);
  });

  test("accepts p:ph type=title as a title placeholder", async () => {
    const presentation = await Presentation.open(rewriteSlide((xml) => xml.replace('type="ctrTitle"', 'type="title"')));
    const slide = required(presentation.slides[0], "slide 1");

    const title = findPlaceholderText(slide, "title");

    expect(title.text).toBe("Original title");
    slide.replaceTextAt(title.anchor, title.text, "Typed title");
    expect(slide.inspectText("unit.placeholders.title-type").map((paragraph) => paragraph.text)).toEqual([
      "Typed title",
      "Original subtitle",
    ]);
  });

  test("uses the direct placeholder shape rather than the first text paragraph", async () => {
    const presentation = await Presentation.open(rewriteSlide((xml) => xml.replace(TITLE_SHAPE, PLAIN_TEXT_SHAPE + TITLE_SHAPE)));
    const slide = required(presentation.slides[0], "slide 1");

    const title = findPlaceholderText(slide, "title");

    expect(title.text).toBe("Original title");
    expect(title.anchor.paragraphIndex).toBe(1);
    slide.replaceTextAt(title.anchor, title.text, "Resolved title");
    expect(slide.inspectText("unit.placeholders.non-title-first").map((paragraph) => paragraph.text)).toEqual([
      "Lead in",
      "Resolved title",
      "Original subtitle",
    ]);
  });

  test("does not resolve nested group placeholders as slide titles", async () => {
    const presentation = await Presentation.open(rewriteSlide((xml) => xml
      .replace(TITLE_SHAPE, TITLE_SHAPE.replace('<p:ph type="ctrTitle" />', ""))
      .replace(TITLE_SHAPE.replace('<p:ph type="ctrTitle" />', ""), NESTED_TITLE_GROUP + TITLE_SHAPE.replace('<p:ph type="ctrTitle" />', ""))));
    const slide = required(presentation.slides[0], "slide 1");

    const error = captureError(() => findPlaceholderText(slide, "title"));

    expect(error.code).toBe("PPTX_PLACEHOLDER_NOT_FOUND");
  });

  test("refuses duplicate direct placeholder shapes", async () => {
    const duplicateTitleShape = TITLE_SHAPE
      .replace('id="2"', 'id="20"')
      .replace('name="Title 1"', 'name="Title duplicate"')
      .replace("Original title", "Duplicate title");
    const presentation = await Presentation.open(rewriteSlide((xml) => xml.replace(TITLE_SHAPE, TITLE_SHAPE + duplicateTitleShape)));
    const slide = required(presentation.slides[0], "slide 1");

    const error = captureError(() => findPlaceholderText(slide, "title"));

    expect(error.code).toBe("PPTX_PLACEHOLDER_AMBIGUOUS");
  });

  test("refuses multi-paragraph subtitle placeholders", async () => {
    const presentation = await Presentation.open(rewriteSlide((xml) => xml.replace(
      SUBTITLE_PARAGRAPH,
      SUBTITLE_PARAGRAPH + '<a:p><a:r><a:t>Second paragraph</a:t></a:r></a:p>',
    )));
    const slide = required(presentation.slides[0], "slide 1");

    const error = captureError(() => findPlaceholderText(slide, "subtitle"));

    expect(error.code).toBe("PPTX_UNSUPPORTED_TEXT_TOPOLOGY");
  });

  test("refuses placeholder paragraphs with break and field topology", async () => {
    const presentation = await Presentation.open(rewriteSlide((xml) => xml.replace(TITLE_PARAGRAPH, BREAK_FIELD_TITLE_PARAGRAPH)));
    const slide = required(presentation.slides[0], "slide 1");

    const error = captureError(() => findPlaceholderText(slide, "title"));

    expect(error.code).toBe("PPTX_UNSUPPORTED_TEXT_TOPOLOGY");
  });

  test("refuses empty title placeholders until empty replacement is supported", async () => {
    const presentation = await Presentation.open(rewriteSlide((xml) => xml.replace(TITLE_PARAGRAPH, "<a:p></a:p>")));
    const slide = required(presentation.slides[0], "slide 1");

    const error = captureError(() => findPlaceholderText(slide, "title"));

    expect(error.code).toBe("PPTX_UNSUPPORTED_TEXT_TOPOLOGY");
  });
});

async function sharedFixtureBytes(): Promise<Uint8Array> {
  return Uint8Array.from(await Bun.file(SHARED_FIXTURE).bytes());
}

function rewriteSlide(rewrite: (xml: string) => string): Uint8Array {
  const parts = readZip(readFixtureBytesSync());
  const original = decodeRequiredPart(parts, SLIDE_PART);
  const updated = rewrite(original);
  if (updated === original) {
    throw new Error("slide rewrite did not change the fixture");
  }
  parts.set(SLIDE_PART, encoder.encode(updated));
  return writeZip(parts);
}

function readFixtureBytesSync(): Uint8Array {
  return Uint8Array.from(readFileSync(SHARED_FIXTURE));
}

function decodeRequiredPart(parts: ReadonlyMap<string, Uint8Array>, name: string): string {
  const part = parts.get(name);
  if (!part) {
    throw new Error(`missing ZIP member ${name}`);
  }
  return decoder.decode(part);
}

function captureError(operation: () => unknown): OoxmlError {
  try {
    operation();
  } catch (error) {
    if (error instanceof OoxmlError) {
      return error;
    }
    throw error;
  }
  throw new Error("expected OoxmlError");
}

function required<T>(value: T | undefined, label: string): T {
  if (value === undefined) {
    throw new Error(`missing ${label}`);
  }
  return value;
}
