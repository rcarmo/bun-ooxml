import { Document, OpcPackage, Presentation, Workbook } from "../src/index.ts";
import { attribute, elements, parseXml } from "../src/xml/index.ts";
import { readZip } from "../src/opc/zip.ts";

const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SENTINEL_PART = "customXml/preservation-sentinel.xml";
const SENTINEL_REL_ID = "rIdSharedSentinel";
const SENTINEL_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml";
const SENTINEL_NS = "urn:rcarmo:ooxml:acceptance:v1";
const SENTINEL_TEXT = "retain these exact bytes";

export interface FixtureOrigin {
  repository: string;
  revision: string;
  path: string;
  sha256: string;
}

export interface FixtureRecord {
  id: string;
  path: string;
  sha256: string;
  bytes: number;
  origin: FixtureOrigin;
  transformation: string;
  facts: Record<string, unknown>;
  allowedChangedPartsForSuccess: string[];
  mustPreservePayloads: Record<string, string>;
  memberSha256: Record<string, string>;
}

export interface FixtureManifest {
  schemaVersion: number;
  contractRevision: string;
  generator: {
    path: string;
    sha256: string;
    zip: string;
  };
  fixtures: FixtureRecord[];
}

export async function verifyFixture(bytes: Uint8Array, fixture: FixtureRecord): Promise<void> {
  assert(
    sha256Hex(bytes) === fixture.sha256,
    `${fixture.id}: archive sha256 drift`,
  );
  assert(
    bytes.length === fixture.bytes,
    `${fixture.id}: archive byte length drift (${bytes.length} !== ${fixture.bytes})`,
  );

  const parts = readZip(bytes);
  assertExpectedMemberSet(parts, fixture, fixture.id);
  assertOnlyAllowedPayloadDrift(parts, fixture, fixture.id, new Set<string>());
  assertPreservedPayloads(parts, fixture, fixture.id);

  await assertPackagePreconditions(bytes, fixture);
  await assertReadBackFacts(bytes, fixture);
}

export async function assertPreservedOutput(bytes: Uint8Array, fixture: FixtureRecord): Promise<void> {
  const parts = readZip(bytes);
  assertExpectedMemberSet(parts, fixture, fixture.id);
  assertOnlyAllowedPayloadDrift(
    parts,
    fixture,
    fixture.id,
    new Set(fixture.allowedChangedPartsForSuccess),
  );
  assertPreservedPayloads(parts, fixture, fixture.id);

  await assertPackagePreconditions(bytes, fixture);
}

function assertExpectedMemberSet(
  parts: ReadonlyMap<string, Uint8Array>,
  fixture: FixtureRecord,
  label: string,
): void {
  const expected = Object.keys(fixture.memberSha256).sort();
  const actual = [...parts.keys()].sort();
  const actualSet = new Set(actual);
  const expectedSet = new Set(expected);
  const added = actual.filter((name) => !expectedSet.has(name));
  const removed = expected.filter((name) => !actualSet.has(name));

  assert(added.length === 0, `${label}: unexpected added parts: ${added.join(", ")}`);
  assert(removed.length === 0, `${label}: unexpected deleted parts: ${removed.join(", ")}`);
}

function assertOnlyAllowedPayloadDrift(
  parts: ReadonlyMap<string, Uint8Array>,
  fixture: FixtureRecord,
  label: string,
  allowedChanged: ReadonlySet<string>,
): void {
  for (const [name, bytes] of parts) {
    const expected = fixture.memberSha256[name];
    assert(expected !== undefined, `${label}: unexpected unpinned part ${name}`);
    const actual = sha256Hex(bytes);
    if (actual === expected) {
      continue;
    }
    assert(allowedChanged.has(name), `${label}: unexpected payload drift in ${name}`);
  }
}

function assertPreservedPayloads(
  parts: ReadonlyMap<string, Uint8Array>,
  fixture: FixtureRecord,
  label: string,
): void {
  for (const [name, expected] of Object.entries(fixture.mustPreservePayloads)) {
    const actualBytes = parts.get(name);
    assert(actualBytes !== undefined, `${label}: missing required preserved part ${name}`);
    assert(
      sha256Hex(actualBytes) === expected,
      `${label}: preserved payload drift in ${name}`,
    );
  }
}

async function assertPackagePreconditions(bytes: Uint8Array, fixture: FixtureRecord): Promise<void> {
  const pkg = await OpcPackage.open(bytes);
  assertSentinel(pkg, fixture);
  if (fixture.path.endsWith(".xlsx")) {
    await assertXlsxStylesResolve(bytes, fixture);
  }
}

function assertSentinel(pkg: OpcPackage, fixture: FixtureRecord): void {
  const sentinel = pkg.get(SENTINEL_PART);
  assert(sentinel !== undefined, `${fixture.id}: missing sentinel part ${SENTINEL_PART}`);

  const rootRelationship = pkg.relationships("").find((relationship) =>
    !relationship.external
    && relationship.id === SENTINEL_REL_ID
    && relationship.type === SENTINEL_REL_TYPE
    && relationship.resolved === SENTINEL_PART
  );
  assert(rootRelationship !== undefined, `${fixture.id}: missing sentinel root relationship`);

  const document = parseXml(pkg.text(SENTINEL_PART));
  assert(
    document.root.localName === "preservation-sentinel" && document.root.namespaceURI === SENTINEL_NS,
    `${fixture.id}: invalid sentinel root element`,
  );
  assert(document.root.text === SENTINEL_TEXT, `${fixture.id}: sentinel payload text drift`);
}

async function assertReadBackFacts(bytes: Uint8Array, fixture: FixtureRecord): Promise<void> {
  switch (fixture.id) {
    case "present-placeholder.docx": {
      const doc = await Document.open(bytes);
      const bodyText = expectStringFact(fixture, "bodyText");
      const absentText = expectStringFact(fixture, "absentText");
      assert(
        doc.paragraphs.map((paragraph) => paragraph.text).join("\n") === bodyText,
        `${fixture.id}: DOCX body text readback drift`,
      );
      assert(doc.find(absentText).length === 0, `${fixture.id}: DOCX absent text unexpectedly present`);
      return;
    }
    case "title-and-subtitle.pptx": {
      const presentation = await Presentation.open(bytes);
      const slideCount = expectNumberFact(fixture, "slideCount");
      const slide1 = expectObjectFact(fixture, "slide1");
      assert(
        presentation.slides.length === slideCount,
        `${fixture.id}: PPTX slide count readback drift`,
      );
      assert(
        JSON.stringify(presentation.slides[0]?.inspectText("shared-fixtures.verify").map((entry) => entry.text) ?? [])
          === JSON.stringify([expectNestedString(slide1, "title"), expectNestedString(slide1, "subtitle")]),
        `${fixture.id}: PPTX title/subtitle readback drift`,
      );
      return;
    }
    case "default-style.xlsx": {
      const workbook = await Workbook.open(bytes);
      const cellText = expectStringFact(fixture, "activeSheetCellA1");
      const absentSheet = expectStringFact(fixture, "absentSheet");
      const cellXfsCount = expectNumberFact(fixture, "cellXfsCount");
      const firstSheet = workbook.sheetnames[0];
      assert(firstSheet !== undefined, `${fixture.id}: workbook has no worksheets`);
      assert(
        workbook.worksheet(firstSheet).getCell("A1")?.value === cellText,
        `${fixture.id}: XLSX A1 readback drift`,
      );
      assert(
        !workbook.sheetnames.includes(absentSheet),
        `${fixture.id}: XLSX absent sheet unexpectedly present`,
      );
      await assertXlsxStylesResolve(bytes, fixture, cellXfsCount);
      return;
    }
    case "cross-sheet-cache.xlsx": {
      const workbook = await Workbook.open(bytes);
      const inputValue = fixture.facts["Input!A1"];
      const calcFact = fixture.facts["Calc!A1"];
      const calcChainPresent = fixture.facts.calcChainPresent;
      assert(
        workbook.worksheet("Input").getCell("A1")?.value === inputValue,
        `${fixture.id}: XLSX Input!A1 readback drift`,
      );
      assert(isRecord(calcFact), `${fixture.id}: missing Calc!A1 facts`);
      const calc = workbook.worksheet("Calc").getCell("A1");
      assert(calc?.kind === "formula", `${fixture.id}: Calc!A1 is not a formula cell`);
      const expectedFormula = typeof calcFact.formula === "string" && calcFact.formula.startsWith("=")
        ? calcFact.formula.slice(1)
        : calcFact.formula;
      assert(calc.formula === expectedFormula, `${fixture.id}: Calc!A1 formula readback drift`);
      assert(calc.cached === calcFact.cachedValue, `${fixture.id}: Calc!A1 cache readback drift`);
      const hasCalcChain = workbook.package.get("xl/calcChain.xml") !== undefined;
      assert(hasCalcChain === calcChainPresent, `${fixture.id}: calcChain presence drift`);
      await assertXlsxStylesResolve(bytes, fixture);
      return;
    }
    default:
      throw new Error(`${fixture.id}: unsupported shared fixture id`);
  }
}

async function assertXlsxStylesResolve(
  bytes: Uint8Array,
  fixture: FixtureRecord,
  expectedCellXfsCount?: number,
): Promise<void> {
  const workbook = await Workbook.open(bytes);
  const stylesPart = workbook.package.related(workbook.package.mainPart(), "styles");
  assert(stylesPart !== undefined, `${fixture.id}: missing workbook styles part`);

  const styles = parseXml(workbook.package.text(stylesPart));
  assert(
    styles.root.localName === "styleSheet" && styles.root.namespaceURI === S_NS,
    `${fixture.id}: invalid styles root element`,
  );

  const cellXfs = elements(styles.root, "cellXfs", S_NS)[0];
  assert(cellXfs !== undefined, `${fixture.id}: missing cellXfs collection`);
  const xfCount = cellXfs.children.filter((child) => child.localName === "xf" && child.namespaceURI === S_NS).length;
  if (expectedCellXfsCount !== undefined) {
    assert(xfCount === expectedCellXfsCount, `${fixture.id}: cellXfs count drift`);
  }

  for (const sheetName of workbook.sheetnames) {
    const sheetPart = resolveWorksheetPart(workbook.package, sheetName);
    const sheet = parseXml(workbook.package.text(sheetPart));
    for (const cell of elements(sheet.root, "c", S_NS)) {
      // An omitted s attribute means style zero, not absence of a dependency.
      // Dropping the default xf must fail even if every cell omits an explicit s.
      const styleIndex = cell.attributes.s ?? "0";
      assert(/^[0-9]+$/.test(styleIndex), `${fixture.id}: invalid style index ${styleIndex} on ${sheetName}!${cell.attributes.r ?? "?"}`);
      const numeric = Number(styleIndex);
      assert(
        numeric < xfCount,
        `${fixture.id}: unresolved style index ${styleIndex} on ${sheetName}!${cell.attributes.r ?? "?"}`,
      );
    }
  }
}

function resolveWorksheetPart(pkg: OpcPackage, sheetName: string): string {
  const workbookPart = pkg.mainPart();
  const workbook = parseXml(pkg.text(workbookPart));
  const relationships = new Map(
    pkg.relationships(workbookPart)
      .filter((relationship) => !relationship.external)
      .map((relationship) => [relationship.id, relationship.resolved]),
  );

  for (const sheet of elements(workbook.root, "sheet", S_NS)) {
    if (sheet.attributes.name !== sheetName) {
      continue;
    }
    const relationshipId = attribute(sheet, "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
    assert(relationshipId !== undefined, `Workbook sheet ${sheetName} is missing an officeDocument relationship id`);
    const resolved = relationships.get(relationshipId);
    assert(resolved !== undefined, `Workbook sheet ${sheetName} has no resolved worksheet part`);
    return resolved;
  }

  throw new Error(`Workbook sheet ${sheetName} not found`);
}

function expectStringFact(fixture: FixtureRecord, key: string): string {
  const value = fixture.facts[key];
  assert(typeof value === "string", `${fixture.id}: fact ${key} must be a string`);
  return value;
}

function expectNumberFact(fixture: FixtureRecord, key: string): number {
  const value = fixture.facts[key];
  assert(typeof value === "number", `${fixture.id}: fact ${key} must be a number`);
  return value;
}

function expectObjectFact(fixture: FixtureRecord, key: string): Record<string, unknown> {
  const value = fixture.facts[key];
  assert(isRecord(value), `${fixture.id}: fact ${key} must be an object`);
  return value;
}

function expectNestedString(value: Record<string, unknown>, key: string): string {
  const nested = value[key];
  assert(typeof nested === "string", `Nested fact ${key} must be a string`);
  return nested;
}

function sha256Hex(bytes: Uint8Array): string {
  return new Bun.CryptoHasher("sha256").update(bytes).digest("hex");
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) {
    throw new Error(message);
  }
}
