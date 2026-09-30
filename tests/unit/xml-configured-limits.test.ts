import { describe, expect, test } from "bun:test";

import {
  inspectXmlEvents,
  parseXml,
  type XmlEventSink,
  validateXml,
} from "../../src/xml/index.ts";

describe("configured XML parser limits", () => {
  test("parseXml accepts the exact configured depth and refuses the next level", () => {
    const exact = `${"<n>".repeat(4)}${"</n>".repeat(4)}`;
    const tooDeep = `${"<n>".repeat(5)}${"</n>".repeat(5)}`;

    expect(parseXml(exact, { maxDepth: 4 }).root.localName).toBe("n");
    expect(() => parseXml(tooDeep, { maxDepth: 4 })).toThrow(
      expect.objectContaining({ code: "XML_DEPTH_LIMIT" }),
    );
  });

  test("inspectXmlEvents accepts the exact configured node count and refuses one more element", () => {
    const seen: string[] = [];
    const sink: XmlEventSink = {
      start(node) { seen.push(node.name); },
      end() {},
      text() {},
      comment() {},
      instruction() {},
    };

    inspectXmlEvents(`<r>${"<n/>".repeat(5)}</r>`, sink, { maxNodes: 6 });
    expect(seen).toEqual(["r", "n", "n", "n", "n", "n"]);
    expect(() => inspectXmlEvents(`<r>${"<n/>".repeat(6)}</r>`, sink, { maxNodes: 6 })).toThrow(
      expect.objectContaining({ code: "XML_NODE_LIMIT" }),
    );
  });

  test("validateXml accepts the exact configured source length and counts UTF-16 code units", () => {
    const exact = `<r>${"x".repeat(57)}</r>`;
    const tooLong = `<r>${"x".repeat(58)}</r>`;

    expect(() => validateXml(exact, { maxSourceUnits: 64 })).not.toThrow();
    expect(() => validateXml(tooLong, { maxSourceUnits: 64 })).toThrow(
      expect.objectContaining({ code: "XML_INPUT_TOO_LARGE" }),
    );
    expect(() => validateXml("<r>π</r>", { maxSourceUnits: 8 })).not.toThrow();
    expect(() => validateXml("<r>😀</r>", { maxSourceUnits: 8 })).toThrow(
      expect.objectContaining({ code: "XML_INPUT_TOO_LARGE" }),
    );
    expect(() => validateXml(`<r>${"x".repeat(57)}\u0001</r>`, { maxSourceUnits: 64 })).toThrow(
      expect.objectContaining({ code: "XML_INPUT_TOO_LARGE" }),
    );
  });

  test("rejects invalid configured limits", () => {
    const sink: XmlEventSink = {
      start() {},
      end() {},
      text() {},
      comment() {},
      instruction() {},
    };

    const cases: Array<() => void> = [
      () => parseXml("<r/>", { maxDepth: 0 }),
      () => parseXml("<r/>", { maxDepth: -1 }),
      () => parseXml("<r/>", { maxDepth: 1.5 }),
      () => validateXml("<r/>", { maxNodes: Number.NaN }),
      () => validateXml("<r/>", { maxNodes: Number.POSITIVE_INFINITY }),
      () => inspectXmlEvents("<r/>", sink, { maxSourceUnits: 0 }),
      () => inspectXmlEvents("<r/>", sink, { maxSourceUnits: 1.1 }),
      () => parseXml("<r/>", null as unknown as never),
    ];

    for (const run of cases) {
      expect(run).toThrow(expect.objectContaining({ code: "XML_LIMIT_INVALID" }));
    }
  });
});
