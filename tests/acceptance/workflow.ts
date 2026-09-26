import { expect } from "bun:test";
import { lstat, mkdtemp, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { basename, join, resolve } from "node:path";

import type { StepBinding } from "../../scripts/gherkin.ts";
import { assertPreservedOutput, verifyFixture, type FixtureRecord } from "../../scripts/shared-fixtures.ts";
import { parseBatchTable } from "../../scripts/shared-pack.ts";
import { Document, OpcPackage, Presentation, Workbook } from "../../src/index.ts";
import { findPlaceholderText } from "../../src/pptx/placeholders.ts";
import { patchOffice, type PatchReceipt } from "../../src/workflow/index.ts";
import { elements, parseXml } from "../../src/xml/index.ts";

const PROJECT_ROOT = resolve(import.meta.dir, "../..");
const PACK_ROOT = join(PROJECT_ROOT, "docs/contracts/shared-v2/pack");
const S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const encoder = new TextEncoder();
const tempRoots: string[] = [];

let manifestPromise: Promise<{ fixtures: FixtureRecord[] }> | undefined;

type DestinationState = "source" | "distinct-absent" | "distinct-existing";
type FileSnapshot = { exists: boolean; bytes?: Uint8Array; sha256: string | null };
type RecordedState = { source: FileSnapshot; destination: FileSnapshot; entries: string[] };
type WorkflowState = {
  fixture?: FixtureRecord;
  tempRoot?: string;
  sourcePath?: string;
  destinationPath?: string;
  requestOutput?: string;
  destinationState?: DestinationState;
  recorded?: RecordedState;
  calculationPolicy?: "invalidate-without-recalculation";
  receipt?: PatchReceipt;
};

export async function cleanupWorkflowFixtures(): Promise<void> {
  await Promise.all(tempRoots.splice(0).map((root) => rm(root, { recursive: true, force: true })));
}

export const bindings: StepBinding[] = [
  {
    pattern: /^fixture "([^"]+)" verified against the fixture manifest$/,
    run: async (context, id) => {
      const state = workflowState(context);
      const fixture = await fixtureById(id);
      const bytes = await Bun.file(join(PACK_ROOT, fixture.path)).bytes();
      await verifyFixture(bytes, fixture);
      const tempRoot = await mkdtemp(join(tmpdir(), "bun-ooxml-workflow-"));
      tempRoots.push(tempRoot);
      const sourcePath = join(tempRoot, fixture.id);
      await Bun.write(sourcePath, bytes);
      Object.assign(state, {
        fixture,
        tempRoot,
        sourcePath,
        destinationPath: join(tempRoot, `output-${basename(fixture.id)}`),
        requestOutput: undefined,
        destinationState: undefined,
        recorded: undefined,
        calculationPolicy: undefined,
        receipt: undefined,
      } satisfies WorkflowState);
    },
  },
  {
    pattern: /^destination state is "(source|distinct-absent|distinct-existing)"$/,
    run: async (context, destinationState) => {
      const state = workflowState(context);
      const sourcePath = required(state.sourcePath, caseLabel(context, "missing source path"));
      const distinctPath = required(state.destinationPath, caseLabel(context, "missing destination path"));
      state.destinationState = destinationState as DestinationState;
      if (destinationState === "source") {
        await rm(distinctPath, { force: true });
        state.requestOutput = undefined;
        return;
      }
      state.requestOutput = distinctPath;
      if (destinationState === "distinct-existing") {
        await Bun.write(distinctPath, encoder.encode(`existing:${caseLabel(context, required(state.fixture, "fixture").id)}`));
        return;
      }
      await rm(distinctPath, { force: true });
      expect(await exists(distinctPath)).toBe(false);
      expect(sourcePath).not.toBe(distinctPath);
    },
  },
  {
    pattern: /^source bytes, destination bytes and document-directory entries are recorded$/,
    run: async (context) => {
      const state = workflowState(context);
      state.recorded = await recordState(state, true);
    },
  },
  {
    pattern: /^source bytes are recorded$/,
    run: async (context) => {
      const state = workflowState(context);
      state.recorded = await recordState(state, false);
    },
  },
  {
    pattern: /^calculation policy is "(invalidate-without-recalculation)"$/,
    run: (context, policy) => {
      workflowState(context).calculationPolicy = policy as WorkflowState["calculationPolicy"];
    },
  },
  {
    pattern: /^previewing this batch without committing:$/,
    run: async (context) => {
      const state = workflowState(context);
      const receipt = await runWorkflow(context, { mode: "dry_run" });
      check(receipt.status === "preview", caseLabel(context, `expected preview receipt, got ${receipt.status}`));
      check(receipt.error === undefined, caseLabel(context, `unexpected preview error: ${receipt.error?.code}`));
      check(receipt.outputSha256 === undefined, caseLabel(context, "preview must not report outputSha256"));
      check(receipt.results.every((result) => result.status === "matched" && result.matched === 1), caseLabel(context, "preview must report matched targets without committed statuses"));
      check(receipt.committedChanges === 0, caseLabel(context, "preview must not commit changes"));
      state.receipt = receipt;
    },
  },
  {
    pattern: /^committing this batch with all-targets-required policy:$/,
    run: async (context) => {
      const receipt = await runWorkflow(context, { mode: "strict" });
      check(receipt.status === "refused", caseLabel(context, `expected refusal, got ${receipt.status}`));
      check(receipt.outputSha256 === undefined, caseLabel(context, "refused receipt must not report outputSha256"));
      check(receipt.changedParts.length === 0, caseLabel(context, "refused receipt must not report changed parts"));
      check(receipt.results.some((result) => result.status === "unmatched"), caseLabel(context, "strict refusal must report an unmatched target"));
      check(receipt.results.every((result) => result.status !== "committed"), caseLabel(context, "refused batch must not report committed targets"));
    },
  },
  {
    pattern: /^committing this batch to the distinct destination:$/,
    run: async (context) => {
      const receipt = await runWorkflow(context, { mode: "safe" });
      check(receipt.status === "committed", caseLabel(context, `expected committed receipt, got ${receipt.status}`));
      check(receipt.error === undefined, caseLabel(context, `unexpected commit error: ${receipt.error?.code}`));
      check(receipt.results.every((result) => result.status === "committed"), caseLabel(context, "successful safe batch must report committed targets"));
      check(receipt.outputSha256 !== undefined, caseLabel(context, "successful safe batch must report outputSha256"));
    },
  },
  {
    pattern: /^committing this batch with multiline wrap enabled:$/,
    run: async (context) => {
      const receipt = await runWorkflow(context, { mode: "safe", multilineWrap: true });
      check(receipt.status === "committed", caseLabel(context, `expected committed receipt, got ${receipt.status}`));
      check(receipt.changedParts.includes("xl/styles.xml"), caseLabel(context, "multiline XLSX commit must report xl/styles.xml in changedParts"));
    },
  },
  {
    pattern: /^committed change count is (\d+)$/,
    run: (context, countText) => {
      expect(required(workflowState(context).receipt, caseLabel(context, "missing receipt")).committedChanges).toBe(Number(countText));
    },
  },
  {
    pattern: /^source bytes, destination bytes and document-directory entries equal the recorded state$/,
    run: async (context) => {
      const state = workflowState(context);
      const recorded = required(state.recorded, caseLabel(context, "missing recorded state"));
      expect(await snapshot(required(state.sourcePath, "sourcePath"))).toEqual(recorded.source);
      expect(await snapshot(logicalDestinationPath(state))).toEqual(recorded.destination);
      expect(await directoryEntries(required(state.tempRoot, "tempRoot"))).toEqual(recorded.entries);
    },
  },
  {
    pattern: /^source bytes equal the recorded state$/,
    run: async (context) => {
      const state = workflowState(context);
      const recorded = required(state.recorded, caseLabel(context, "missing recorded state"));
      expect(await snapshot(required(state.sourcePath, "sourcePath"))).toEqual(recorded.source);
    },
  },
  {
    pattern: /^the operation is refused for unmatched target "([^"]+)"$/,
    run: (context, target) => {
      const receipt = required(workflowState(context).receipt, caseLabel(context, "missing receipt"));
      check(receipt.status === "refused", caseLabel(context, `expected refusal for ${target}`));
      check(receipt.error?.code === "workflow-target-missing", caseLabel(context, `expected workflow-target-missing for ${target}, got ${receipt.error?.code}`));
      const result = receipt.results.find((entry) => entry.target === target);
      check(result?.status === "unmatched" && result.matched === 0 && result.code === "workflow-target-missing", caseLabel(context, `missing unmatched per-target result for ${target}`));
      check(receipt.results.every((entry) => entry.status !== "committed"), caseLabel(context, `refused receipt incorrectly reports committed target for ${target}`));
    },
  },
  {
    pattern: /^"([^"]+)" is never reported applied$/,
    run: (context, target) => {
      const receipt = required(workflowState(context).receipt, caseLabel(context, "missing receipt"));
      const result = receipt.results.find((entry) => entry.target === target);
      check(result !== undefined, caseLabel(context, `missing per-target result for ${target}`));
      check(result.status !== "committed" && result.status !== "unchanged", caseLabel(context, `${target} was incorrectly reported applied as ${result.status}`));
    },
  },
  {
    pattern: /^the reopened source slide 1 title is "([^"]+)"$/,
    run: async (context, expected) => {
      const presentation = await Presentation.open(required(workflowState(context).sourcePath, "sourcePath"));
      expect(findPlaceholderText(required(presentation.slides[0], "slide 1"), "title").text).toBe(expected);
    },
  },
  {
    pattern: /^the reopened source current body text is "([^"]+)"$/,
    run: async (context, expected) => {
      const document = await Document.open(required(workflowState(context).sourcePath, "sourcePath"));
      expect(document.paragraphs.map((paragraph) => paragraph.text).join("\n")).toBe(expected);
    },
  },
  {
    pattern: /^the reopened source active sheet cell A1 is "([^"]+)"$/,
    run: async (context, expected) => {
      const workbook = await Workbook.open(required(workflowState(context).sourcePath, "sourcePath"));
      const firstSheet = required(workbook.sheetnames[0], caseLabel(context, "workbook has no sheets"));
      expect(workbook.worksheet(firstSheet).getCell("A1")?.value).toBe(expected);
    },
  },
  {
    pattern: /^the reopened destination slide 1 title is "([^"]+)"$/,
    run: async (context, expected) => {
      const presentation = await Presentation.open(destinationPath(workflowState(context)));
      expect(findPlaceholderText(required(presentation.slides[0], "slide 1"), "title").text).toBe(expected);
    },
  },
  {
    pattern: /^the reopened destination slide 1 subtitle is "([^"]+)"$/,
    run: async (context, expected) => {
      const presentation = await Presentation.open(destinationPath(workflowState(context)));
      expect(findPlaceholderText(required(presentation.slides[0], "slide 1"), "subtitle").text).toBe(expected);
    },
  },
  {
    pattern: /^the reopened destination active sheet cell A1 has value_json "([^"]+)"$/,
    run: async (context, expectedJsonText) => {
      const workbook = await Workbook.open(destinationPath(workflowState(context)));
      const firstSheet = required(workbook.sheetnames[0], caseLabel(context, "workbook has no sheets"));
      expect(workbook.worksheet(firstSheet).getCell("A1")?.value).toBe(JSON.parse(`"${expectedJsonText}"`));
    },
  },
  {
    pattern: /^that cell resolves to a style with wrapText enabled$/,
    run: async (context) => {
      const workbook = await Workbook.open(destinationPath(workflowState(context)));
      expect(cellWrapTextEnabled(workbook, required(workbook.sheetnames[0], "sheet"), "A1")).toBe(true);
      expect(required(workflowState(context).receipt, caseLabel(context, "missing receipt")).changedParts).toContain("xl/styles.xml");
    },
  },
  {
    pattern: /^every cell style index is below the output cellXfs count$/,
    run: async (context) => {
      const workbook = await Workbook.open(destinationPath(workflowState(context)));
      expect(allStyleIndicesResolve(workbook)).toBe(true);
    },
  },
  {
    pattern: /^the reopened destination Input!A1 is numeric (\d+)$/,
    run: async (context, numericText) => {
      const workbook = await Workbook.open(destinationPath(workflowState(context)));
      expect(workbook.worksheet("Input").getCell("A1")).toMatchObject({ kind: "number", value: Number(numericText) });
    },
  },
  {
    pattern: /^the reopened destination Calc!A1 formula is "([^"]+)"$/,
    run: async (context, expected) => {
      const workbook = await Workbook.open(destinationPath(workflowState(context)));
      const cell = workbook.worksheet("Calc").getCell("A1");
      expect(cell).toMatchObject({ kind: "formula", formula: expected.startsWith("=") ? expected.slice(1) : expected });
      expect(required(workflowState(context).receipt, caseLabel(context, "missing receipt")).changedParts.sort()).toEqual([
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/sheet2.xml",
      ]);
    },
  },
  {
    pattern: /^the destination Calc!A1 cached value is absent or empty$/,
    run: async (context) => {
      const workbook = await Workbook.open(destinationPath(workflowState(context)));
      const cell = workbook.worksheet("Calc").getCell("A1");
      check(cell?.kind === "formula", caseLabel(context, "Calc!A1 is not a formula cell"));
      check(cell.cached === null, caseLabel(context, `expected empty cached formula value, got ${String(cell.cached)}`));
      const cached = rawCellValue(workbook, "Calc", "A1");
      check(cached === undefined || cached === "", caseLabel(context, `expected absent or empty raw cached value, got ${JSON.stringify(cached)}`));
    },
  },
  {
    pattern: /^calculation state is "([^"]+)"$/,
    run: (context, expected) => {
      expect(String(required(workflowState(context).receipt, caseLabel(context, "missing receipt")).calculationState)).toBe(expected);
    },
  },
  {
    pattern: /^a data-only read never returns the old cached value (\d+) as current$/,
    run: async (context, oldValueText) => {
      const workbook = await Workbook.open(destinationPath(workflowState(context)));
      const cell = workbook.worksheet("Calc").getCell("A1");
      check(cell?.kind === "formula", caseLabel(context, "Calc!A1 stopped reading as a formula cell"));
      check(cell.cached !== Number(oldValueText), caseLabel(context, `formula cache still exposes stale value ${oldValueText}`));
      check(rawCellValue(workbook, "Calc", "A1") !== oldValueText, caseLabel(context, `raw worksheet cache still contains stale value ${oldValueText}`));
    },
  },
  {
    pattern: /^all destination relationship and content-type references resolve$/,
    run: async (context) => {
      const bytes = await Bun.file(destinationPath(workflowState(context))).bytes();
      const pkg = await OpcPackage.open(bytes);
      expect(pkg.names().length).toBeGreaterThan(0);
    },
  },
  {
    pattern: /^destination member payloads outside the manifest change allowance are byte-identical$/,
    run: async (context) => {
      const state = workflowState(context);
      await assertPreservedOutput(
        await Bun.file(destinationPath(state)).bytes(),
        required(state.fixture, caseLabel(context, "missing fixture metadata")),
      );
    },
  },
];

async function runWorkflow(
  context: Record<string, unknown>,
  options: { mode: "dry_run" | "strict" | "safe"; multilineWrap?: boolean },
): Promise<PatchReceipt> {
  const state = workflowState(context);
  const sourcePath = required(state.sourcePath, caseLabel(context, "missing source path"));
  const requestOutput = state.requestOutput;
  const changes = parseBatchTable(requiredDataTable(context));
  const receipt = await patchOffice({
    source: sourcePath,
    output: requestOutput,
    mode: options.mode,
    changes,
    multilineWrap: options.multilineWrap,
    calculationPolicy: state.calculationPolicy,
  });
  const sourceBytes = await Bun.file(sourcePath).bytes();
  expect(receipt.results.map(({ target, value }) => ({ target, value }))).toEqual(changes);
  expect(receipt.sourceSha256).toBe(sha256(sourceBytes));
  if (receipt.status === "committed") {
    const outputBytes = await Bun.file(destinationPath(state)).bytes();
    expect(receipt.outputSha256).toBe(sha256(outputBytes));
  } else {
    expect(receipt.outputSha256).toBeUndefined();
  }
  state.receipt = receipt;
  return receipt;
}

function workflowState(context: Record<string, unknown>): WorkflowState {
  return context.state as WorkflowState;
}

function requiredDataTable(context: Record<string, unknown>): string[][] {
  const table = (context.step as { argument?: { dataTable?: string[][] } } | undefined)?.argument?.dataTable;
  return required(table, caseLabel(context, "missing data table argument"));
}

async function recordState(state: WorkflowState, includeDestination: boolean): Promise<RecordedState> {
  const sourcePath = required(state.sourcePath, "missing source path");
  const tempRoot = required(state.tempRoot, "missing temp root");
  return {
    source: await snapshot(sourcePath),
    destination: includeDestination ? await snapshot(logicalDestinationPath(state)) : { exists: false, sha256: null },
    entries: await directoryEntries(tempRoot),
  };
}

function logicalDestinationPath(state: WorkflowState): string {
  return state.destinationState === "source"
    ? required(state.sourcePath, "missing source path")
    : required(state.destinationPath, "missing destination path");
}

function destinationPath(state: WorkflowState): string {
  const path = logicalDestinationPath(state);
  return path;
}

async function snapshot(path: string): Promise<FileSnapshot> {
  if (!(await exists(path))) return { exists: false, sha256: null };
  const bytes = await Bun.file(path).bytes();
  return { exists: true, bytes, sha256: sha256(bytes) };
}

async function exists(path: string): Promise<boolean> {
  try {
    await lstat(path);
    return true;
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code === "ENOENT") return false;
    throw error;
  }
}

async function directoryEntries(root: string): Promise<string[]> {
  return (await readdir(root)).sort();
}

async function fixtureById(id: string): Promise<FixtureRecord> {
  manifestPromise ??= Bun.file(join(PACK_ROOT, "fixture-manifest.json")).json() as Promise<{ fixtures: FixtureRecord[] }>;
  const fixture = (await manifestPromise).fixtures.find((entry) => entry.id === id);
  return required(fixture, `missing shared fixture ${id}`);
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
    if (sheet.attributes.name !== sheetName) continue;
    return required(relationships.get(required(sheet.attributes["r:id"], `missing r:id for ${sheetName}`)), `missing worksheet part for ${sheetName}`);
  }
  throw new Error(`missing worksheet ${sheetName}`);
}

function cellWrapTextEnabled(workbook: Workbook, sheetName: string, address: string): boolean {
  const cell = workbook.worksheet(sheetName).getCell(address);
  const styleIndex = Number(cell?.styleId ?? "0");
  const stylesPart = required(workbook.package.related(workbook.package.mainPart(), "styles"), "missing styles part");
  const styles = parseXml(workbook.package.text(stylesPart));
  const cellXfs = required(elements(styles.root, "cellXfs", S_NS)[0], "missing cellXfs");
  const xf = required(elements(cellXfs, "xf", S_NS)[styleIndex], `missing xf ${styleIndex}`);
  return elements(xf, "alignment", S_NS)[0]?.attributes.wrapText === "1";
}

function allStyleIndicesResolve(workbook: Workbook): boolean {
  const stylesPart = required(workbook.package.related(workbook.package.mainPart(), "styles"), "missing styles part");
  const styles = parseXml(workbook.package.text(stylesPart));
  const xfCount = elements(required(elements(styles.root, "cellXfs", S_NS)[0], "missing cellXfs"), "xf", S_NS).length;
  return workbook.sheetnames.every((sheetName) => {
    const sheet = parseXml(workbook.package.text(resolveWorksheetPart(workbook, sheetName)));
    return elements(sheet.root, "c", S_NS).every((cell) => Number(cell.attributes.s ?? "0") < xfCount);
  });
}

function rawCellValue(workbook: Workbook, sheetName: string, address: string): string | undefined {
  const sheet = parseXml(workbook.package.text(resolveWorksheetPart(workbook, sheetName)));
  const cell = elements(sheet.root, "c", S_NS).find((entry) => entry.attributes.r?.toUpperCase() === address.toUpperCase());
  return elements(required(cell, `missing cell ${sheetName}!${address}`), "v", S_NS)[0]?.text;
}

function sha256(bytes: Uint8Array): string {
  return new Bun.CryptoHasher("sha256").update(bytes).digest("hex");
}

function caseLabel(context: Record<string, unknown>, message: string): string {
  const acceptanceCase = context.case as { id?: string; name?: string } | undefined;
  return `${acceptanceCase?.id ?? acceptanceCase?.name ?? "workflow case"}: ${message}`;
}

function required<T>(value: T | undefined, message: string): T {
  if (value === undefined) throw new Error(message);
  return value;
}

function check(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

export default bindings;
