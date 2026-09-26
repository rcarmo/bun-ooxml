import { describe, expect, test } from "bun:test";

import {
  applyEdits,
  elements,
  escapeAttribute,
  escapeText,
  parseXml,
} from "../../src/xml/index.ts";
import { OoxmlError } from "../../src/errors.ts";

const XML_NS = "http://www.w3.org/XML/1998/namespace";

describe("parseXml", () => {
  test("parses namespaces, mixed content and UTF-16 offsets", () => {
    const xml =
      '<?xml version="1.0"?><!--before--><?pi ok?><ns:root xmlns="urn:default" xmlns:ns="urn:ns" xmlns:x="urn:x" a="1 &amp; 2">pre<x:child x:b="v"/>mid<![CDATA[<tail>]]></ns:root><!--after-->';

    const doc = parseXml(xml);
    const root = doc.root;
    const child = root.children[0];

    expect(doc.elements).toHaveLength(2);
    expect(root.name).toBe("ns:root");
    expect(root.localName).toBe("root");
    expect(root.namespaceURI).toBe("urn:ns");
    expect(root.attributes).toEqual({
      xmlns: "urn:default",
      "xmlns:ns": "urn:ns",
      "xmlns:x": "urn:x",
      a: "1 & 2",
    });
    expect(root.text).toBe("premid<tail>");
    expect(root.start).toBe(xml.indexOf("<ns:root"));
    expect(root.openEnd).toBe(xml.indexOf(">", root.start) + 1);
    expect(root.closeStart).toBe(xml.lastIndexOf("</ns:root>"));
    expect(root.end).toBe(root.closeStart + "</ns:root>".length);
    expect(root.selfClosing).toBeFalse();
    expect(root.parent).toBeUndefined();
    expect(root.root).toBe(root);

    expect(child).toBeDefined();
    expect(child?.name).toBe("x:child");
    expect(child?.localName).toBe("child");
    expect(child?.namespaceURI).toBe("urn:x");
    expect(child?.attributes).toEqual({ "x:b": "v" });
    expect(child?.text).toBe("");
    expect(child?.parent).toBe(root);
    expect(child?.root).toBe(root);
    expect(child?.selfClosing).toBeTrue();
    expect(child?.start).toBe(xml.indexOf("<x:child"));
    expect(child?.openEnd).toBe(xml.indexOf("/>", child!.start) + 2);
    expect(child?.closeStart).toBe(child?.openEnd);
    expect(child?.end).toBe(child?.openEnd);

    expect(elements(doc, "child", "urn:x")).toEqual([child!]);
    expect(elements(root, "root", "urn:ns")).toEqual([root]);
  });

  test("accepts predefined and numeric entities", () => {
    const doc = parseXml('<r a="&quot;&apos;">&#x41;&#65;&amp;&lt;&gt;</r>');
    expect(doc.root.attributes).toEqual({ a: '"\'' });
    expect(doc.root.text).toBe("AA&<>");
  });

  test("accepts xml-stylesheet processing instructions before the root element", () => {
    const doc = parseXml('<?xml-stylesheet href="style.xsl"?><r/>');
    expect(doc.root.name).toBe("r");
  });

  test("normalises literal XML whitespace but preserves character references and original offsets", () => {
    const source = '<r a="x\r\ny\tz&#xD;&#xA;&#x9;">u\r\nv\rw&#xD;<![CDATA[c\r\nd]]><s/></r>';
    const doc = parseXml(source);
    expect(doc.root.text).toBe('u\nv\nw\rc\nd');
    expect(doc.root.attributes.a).toBe('x y z\r\n\t');
    const value='x\r\n\ty';
    expect(parseXml(`<r a="${escapeAttribute(value)}">${escapeText(value)}</r>`).root.attributes.a).toBe(value);
    expect(parseXml(`<r>${escapeText(value)}</r>`).root.text).toBe(value);
    const child = doc.root.children[0]!;
    expect(source.slice(child.start,child.end)).toBe('<s/>');
  });

  test("keeps the xml prefix bound implicitly", () => {
    const doc = parseXml('<r xml:lang="en"/>');
    expect(doc.root.attributes).toEqual({ "xml:lang": "en" });
    expect(doc.root.namespaceURI).toBe("");
    expect(XML_NS).toBe("http://www.w3.org/XML/1998/namespace");
  });

  test("stores attributes in null-prototype maps", () => {
    const doc = parseXml('<r __proto__="polluted" constructor="safe"/>');
    expect(Object.getPrototypeOf(doc.root.attributes)).toBeNull();
    expect(Object.prototype.hasOwnProperty.call(doc.root.attributes, "__proto__")).toBeTrue();
    expect(doc.root.attributes["__proto__"]).toBe("polluted");
    expect(doc.root.attributes["constructor"]).toBe("safe");
  });

  test("rejects malformed or unsafe XML", () => {
    const cases: Array<[string, string]> = [
      ['<?xml version="1.0" extra="x"?><r/>', "XML_MALFORMED"],
      ['<?xml version="1.0" standalone="maybe"?><r/>', "XML_MALFORMED"],
      ["<?pi/?>", "XML_MALFORMED"],
      ["<!DOCTYPE r><r/>", "XML_DTD_FORBIDDEN"],
      ["<r>&custom;</r>", "XML_ENTITY_FORBIDDEN"],
      ["<r a='1' a='2'/>", "XML_DUPLICATE_ATTRIBUTE"],
      ["<r a='1'b='2'/>", "XML_MALFORMED"],
      ['<r xmlns:x="u" xmlns:y="u" x:a="1" y:a="2"/>', "XML_DUPLICATE_ATTRIBUTE"],
      ["<xmlns:r/>", "XML_MALFORMED"],
      ['<r xmlns="http://www.w3.org/XML/1998/namespace"/>', "XML_MALFORMED"],
      ["<x:r/>", "XML_UNBOUND_PREFIX"],
      ["<a></b>", "XML_MISMATCHED_TAG"],
      ["<r>\u0001</r>", "XML_INVALID_CHAR"],
      ["<r><!-- bad -- --></r>", "XML_MALFORMED"],
      ["<r><!--bad---></r>", "XML_MALFORMED"],
    ];

    for (const [xml, code] of cases) {
      expect(() => parseXml(xml)).toThrow(
        expect.objectContaining({
          code,
        }),
      );
    }
  });

  test("refuses excessive depth, node count and input length", () => {
    const deep = `${"<n>".repeat(257)}${"</n>".repeat(257)}`;
    const wide = `<r>${"<n/>".repeat(100_000)}</r>`;
    const huge = `<r>${"a".repeat(8 * 1024 * 1024)}</r>`;

    expect(() => parseXml(deep)).toThrow(
      expect.objectContaining({ code: "XML_DEPTH_LIMIT" }),
    );
    expect(() => parseXml(wide)).toThrow(
      expect.objectContaining({ code: "XML_NODE_LIMIT" }),
    );
    expect(() => parseXml(huge)).toThrow(
      expect.objectContaining({ code: "XML_INPUT_TOO_LARGE" }),
    );
  });
});

describe("escaping helpers", () => {
  test("escape text and attribute values", () => {
    expect(escapeText(`5 < 7 & 9 > 4`)).toBe("5 &lt; 7 &amp; 9 &gt; 4");
    expect(escapeAttribute(`'\"<&>`)).toBe("&apos;&quot;&lt;&amp;&gt;");
    expect(() => escapeText("\u0001")).toThrow(
      expect.objectContaining({ code: "XML_INVALID_CHAR" }),
    );
  });
});

describe("applyEdits", () => {
  test("applies sorted disjoint edits with escaped text and XML fragments", () => {
    const xml = "<r>one two</r>";
    const result = applyEdits(xml, [
      {
        start: xml.indexOf("two"),
        end: xml.indexOf("two") + 3,
        value: "<x/>",
      },
      {
        start: xml.indexOf("one"),
        end: xml.indexOf("one") + 3,
        value: escapeText("1 < 2"),
      },
    ]);

    expect(result).toBe("<r>1 &lt; 2 <x/></r>");
    const doc = parseXml(result);
    expect(doc.root.text).toBe("1 < 2 ");
    expect(doc.root.children.map((child) => child.name)).toEqual(["x"]);
  });

  test("refuses overlapping bounds and unsafe resulting XML", () => {
    const xml = "<r>text</r>";

    expect(() =>
      applyEdits(xml, [
        { start: 3, end: 5, value: "a" },
        { start: 4, end: 6, value: "b" },
      ]),
    ).toThrow(expect.objectContaining({ code: "XML_EDIT_OVERLAP" }));

    expect(() =>
      applyEdits(xml, [
        { start: xml.indexOf("text"), end: xml.indexOf("text") + 4, value: "<!DOCTYPE x><x/>" },
      ]),
    ).toThrow(expect.objectContaining({ code: "XML_EDIT_UNSAFE" }));
  });

  test("returns OoxmlError instances", () => {
    expect(() => parseXml("<a></b>")).toThrow(OoxmlError);
  });
});
