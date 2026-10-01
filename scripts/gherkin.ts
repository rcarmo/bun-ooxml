import {
  AstBuilder,
  compile,
  GherkinClassicTokenMatcher,
  Parser,
} from "@cucumber/gherkin";
import {
  IdGenerator,
  PickleStepType,
  type Examples,
  type Feature,
  type GherkinDocument,
  type Location,
  type Pickle,
  type PickleStep,
  type PickleStepArgument,
  type Rule,
  type Scenario,
  type Step,
  type TableRow,
  type Tag,
} from "@cucumber/messages";
import { createHash, randomUUID } from "node:crypto";
import { join, resolve } from "node:path";
import {verifyReferences} from './references.ts';
import pptxManipulationPin from '../docs/behaviors/pptx-manipulation-candidate.json';
import pptxFormattingPin from '../docs/behaviors/pptx-formatting-candidate.json';

export type StepBinding = {
  pattern: RegExp;
  run: (
    context: Record<string, unknown>,
    ...captures: string[]
  ) => unknown | Promise<unknown>;
};

export type AcceptanceLifecycle = "implemented" | "planned";
export type AcceptanceResult = "passed" | "failed" | "planned";
export type AcceptanceStepResult =
  | "passed"
  | "failed"
  | "planned"
  | "skipped"
  | "undefined"
  | "ambiguous";

export type SourceLocation = {
  uri: string;
  line: number;
  column?: number;
};

export type StepArgumentView = {
  docString?: {
    content: string;
    mediaType?: string;
  };
  dataTable?: string[][];
};

export type AcceptanceStep = {
  keyword: string;
  text: string;
  type: PickleStepType;
  source: {
    sha256: string;
    location: SourceLocation;
    exampleLocation?: SourceLocation;
  };
  fromBackground: boolean;
  argument?: StepArgumentView;
};

export type AcceptanceCase = {
  pickleId: string;
  caseId: string;
  identityKey: string;
  name: string;
  location: SourceLocation;
  example?: {
    values: Record<string, string>;
    orderedPairs: Array<{ name: string; value: string }>;
    location: SourceLocation;
  };
  tags: string[];
  steps: AcceptanceStep[];
};

export type AcceptanceScenario = {
  /** Runtime-specific status when a shared feature contains unbound scenarios. */
  lifecycle?: AcceptanceLifecycle;
  scenarioId: string;
  name: string;
  location: SourceLocation;
  tags: string[];
  cases: AcceptanceCase[];
};

export type AcceptanceFeature = {
  path: string;
  name: string;
  lifecycle: AcceptanceLifecycle;
  runner?: "bun";
  /** Canonical source tags. Runtime activation is represented by lifecycle/runner. */
  tags: string[];
  sourceSha256: string;
  scenarios: AcceptanceScenario[];
};

export type InventoryCount = {
  implemented: number;
  planned: number;
  total: number;
};

export type InventoryCoverage = {
  /** Inventory accounting only; this does not assert execution or conformance. */
  catalogue?: { path: string; sourceSha256: string; features: number; scenarios: number; cases: number };
  shared: AcceptanceInventory['counts'];
  localOnly: AcceptanceInventory['counts'];
};

export type AcceptanceInventory = {
  root: string;
  coverage?: InventoryCoverage;
  features: AcceptanceFeature[];
  counts: {
    features: InventoryCount;
    scenarios: InventoryCount;
    cases: InventoryCount;
    steps: InventoryCount;
  };
};

export type AcceptanceStepReport = AcceptanceStep & {
  status: AcceptanceStepResult;
  binding?: string;
  captures?: string[];
  durationMs?: number;
  error?: string;
};

export type AcceptanceCaseReport = Omit<AcceptanceCase, "steps"> & {
  lifecycle: AcceptanceLifecycle;
  result: AcceptanceResult;
  steps: AcceptanceStepReport[];
  error?: string;
};

export type AcceptanceScenarioReport = Omit<AcceptanceScenario, "cases"> & {
  lifecycle: AcceptanceLifecycle;
  result: AcceptanceResult;
  cases: AcceptanceCaseReport[];
};

export type AcceptanceFeatureReport = Omit<AcceptanceFeature, "scenarios"> & {
  result: AcceptanceResult;
  scenarios: AcceptanceScenarioReport[];
};

export type ExecutionCount = {
  passed: number;
  failed: number;
  planned: number;
  total: number;
};

export type StepExecutionCount = {
  passed: number;
  failed: number;
  planned: number;
  skipped: number;
  undefined: number;
  ambiguous: number;
  total: number;
};

export type AcceptanceExecution = {
  features: AcceptanceFeatureReport[];
  counts: {
    features: ExecutionCount;
    scenarios: ExecutionCount;
    cases: ExecutionCount;
    steps: StepExecutionCount;
  };
  failures: string[];
};

type IndexedRow = {
  header: string[];
  values: string[];
  valuesRecord: Record<string, string>;
  location: Location;
};

type IndexedStep = {
  step: Step;
  fromBackground: boolean;
};

type DocumentIndex = {
  steps: Map<string, IndexedStep>;
  rows: Map<string, IndexedRow>;
};

type BindingMatch = {
  binding: StepBinding;
  captures: string[];
};

const FEATURE_GLOB = "features/**/*.feature";
const FEATURE_STATUS_TAGS = ["@implemented", "@planned"] as const;
const FEATURE_RUNNER_TAGS = ["@bun"] as const;
const RESERVED_SCENARIO_TAGS: ReadonlySet<string> = new Set([
  ...FEATURE_STATUS_TAGS,
  ...FEATURE_RUNNER_TAGS,
]);

export function newAcceptanceRunId(): string {
  return randomUUID();
}

export function parseFeature(path: string, text: string, options: { allowDuplicateCaseNames?: boolean } = {}): AcceptanceFeature {
  const ids = IdGenerator.incrementing();
  const document = new Parser(
    new AstBuilder(ids),
    new GherkinClassicTokenMatcher(),
  ).parse(text) as GherkinDocument;
  const feature = document.feature;
  if (!feature) {
    throw new Error(`Missing feature: ${path}`);
  }

  const featureTags = feature.tags.map((tag) => tag.name);
  const lifecycle = validateFeatureTags(path, featureTags);
  const sourceSha256 = sha256(text);
  const index = indexDocument(path, feature);
  const pickles = compile(document, path, ids);
  const scenarios = collectScenarios(feature);

  if (!feature.name.trim()) {
    throw new Error(`Feature name must not be empty: ${path}`);
  }
  if (!scenarios.length) {
    throw new Error(`Feature must contain at least one scenario: ${path}`);
  }

  const scenarioNames = new Set<string>();
  const scenarioIdTags = new Set<string>();
  const parsedScenarios = scenarios.map((scenario) => {
    const scenarioTagNames = scenario.tags.map((tag) => tag.name);
    const scenarioId = validateScenarioTags(path, scenario, scenarioTagNames);
    if (scenarioNames.has(scenario.name.trim())) {
      throw new Error(
        `Duplicate scenario name in ${path}: ${JSON.stringify(scenario.name)}`,
      );
    }
    scenarioNames.add(scenario.name.trim());
    if (scenarioIdTags.has(scenarioId)) {
      throw new Error(`Duplicate scenario id ${scenarioId} in ${path}`);
    }
    scenarioIdTags.add(scenarioId);
    if (!scenario.name.trim()) {
      throw new Error(`Scenario name must not be empty: ${path}:${scenario.location.line}`);
    }
    if (!scenario.steps.length) {
      throw new Error(`Scenario must contain steps: ${path}:${scenario.location.line}`);
    }

    for (const examples of scenario.examples) {
      validateExamplesTags(path, examples);
    }

    const scenarioPickles = pickles.filter((pickle) =>
      pickle.astNodeIds.includes(scenario.id),
    );
    if (!scenarioPickles.length) {
      throw new Error(
        `Scenario has no executable cases: ${path}:${scenario.location.line}`,
      );
    }

    const caseKeys = new Set<string>();
    const caseNames = new Set<string>();
    const cases = scenarioPickles.map((pickle) => {
      const parsedCase = buildCase(path, sourceSha256, index, scenarioId, scenario, pickle);
      if (!parsedCase.name.trim()) {
        throw new Error(`Case name must not be empty: ${path}:${parsedCase.location.line}`);
      }
      if (!options.allowDuplicateCaseNames && caseNames.has(parsedCase.name)) {
        throw new Error(
          `Duplicate expanded case name for ${scenarioId} in ${path}: ${JSON.stringify(parsedCase.name)}`,
        );
      }
      caseNames.add(parsedCase.name);
      if (caseKeys.has(parsedCase.identityKey)) {
        throw new Error(
          `Duplicate expanded case identity for ${scenarioId} in ${path}: ${parsedCase.caseId}`,
        );
      }
      caseKeys.add(parsedCase.identityKey);
      if (!parsedCase.steps.some((step) => step.type === PickleStepType.OUTCOME)) {
        throw new Error(
          `Scenario must include at least one Then step: ${path}:${scenario.location.line}`,
        );
      }
      return parsedCase;
    });

    return {
      scenarioId,
      name: scenario.name.trim(),
      location: toSourceLocation(path, scenario.location),
      tags: scenarioTagNames,
      cases,
    } satisfies AcceptanceScenario;
  });

  return {
    path,
    name: feature.name.trim(),
    lifecycle,
    runner: lifecycle === "implemented" ? "bun" : undefined,
    tags: featureTags,
    sourceSha256,
    scenarios: parsedScenarios,
  } satisfies AcceptanceFeature;
}

export function selectSharedScenarios(path: string, text: string, scenarioIds: unknown): AcceptanceFeature {
  const feature = parseFeature(path, text, {allowDuplicateCaseNames: true});
  if (feature.lifecycle !== 'planned') throw new Error('Shared feature must remain planned: ' + path);
  if (!Array.isArray(scenarioIds) || !scenarioIds.length) throw new Error('Invalid shared scenario selection: empty or missing IDs');
  if (new Set(scenarioIds).size !== scenarioIds.length) throw new Error('Duplicate shared scenario selection');
  const known = new Set(feature.scenarios.map(s => s.scenarioId));
  for (const id of scenarioIds) if (typeof id !== 'string' || !known.has(id)) throw new Error('Unknown shared scenario: ' + String(id));
  const selected = new Set(scenarioIds);
  // Shared outline display names are descriptive, not identities. parseFeature
  // rejects duplicate identity keys; executeAcceptance reports each row separately.
  return {...feature, lifecycle: 'implemented', runner: 'bun', scenarios: feature.scenarios.map(s => ({...s, lifecycle: selected.has(s.scenarioId) ? 'implemented' : 'planned'}))};
}

export async function inventoryFeatures(root: string): Promise<AcceptanceInventory> {
  const features: AcceptanceFeature[] = [];
  for await (const path of new Bun.Glob(FEATURE_GLOB).scan({ cwd: root })) {
    const text = await Bun.file(join(root, path)).text();
    features.push(parseFeature(path, text));
  }
  let catalogue: InventoryCoverage['catalogue'];
  const sharedPrefix='references/fixtures-ooxml/';
  const referenceRoot=process.env.OOXML_FIXTURES_ROOT && resolve(root)===resolve(import.meta.dir,'..')
    ? process.env.OOXML_FIXTURES_ROOT : join(root,sharedPrefix);
  // Candidate admission is an exact clean reference check, not a commit whitelist.
  if(process.env.OOXML_FIXTURES_ROOT&&resolve(root)===resolve(import.meta.dir,'..'))await verifyReferences(root);
  const candidatePin=process.env.OOXML_FIXTURES_ROOT&&resolve(root)===resolve(import.meta.dir,'..')
    ? await Bun.file(process.env.OOXML_REFERENCE_PIN!).json() : undefined;
  const formattingCandidate=candidatePin?.commit===pptxFormattingPin.commit;
  const pptxCandidate=candidatePin?.commit===pptxManipulationPin.commit||formattingCandidate;
  if(pptxCandidate){
    const expected=formattingCandidate?pptxFormattingPin:pptxManipulationPin;
    if(JSON.stringify(candidatePin.selectedScenarioIds)!==JSON.stringify(expected.selectedScenarioIds))throw Error('PPTX candidate selection drift');
    // Migrate the predecessor local lane only for the sealed shared candidate.
    const local=features.findIndex(f=>f.path==='features/pptx/manipulation-next20.feature');
    if(local<0)throw Error('Missing PPTX predecessor lane');features.splice(local,1);
  }
  const sharedConfig = Bun.file(join(root, 'features/shared.json'));
  if (await sharedConfig.exists()) {
    const shared = await sharedConfig.json();
    if (shared.schemaVersion !== 2 || !Array.isArray(shared.features)) throw new Error('Invalid shared feature mapping schema');
    const selections=new Map<string,unknown>();
    const safeFeature=(path:unknown):path is string=>typeof path==='string'&&/^workflows\/(?:[a-z0-9-]+\/)*[a-z0-9-]+\.feature$/.test(path);
    for (const entry of shared.features) {
      if (!entry || typeof entry.path !== 'string' || !entry.path.startsWith(sharedPrefix) || !safeFeature(entry.path.slice(sharedPrefix.length)) || entry.lifecycle !== 'implemented' || entry.runner !== 'bun') throw new Error('Invalid shared feature mapping');
      if(selections.has(entry.path))throw Error('Duplicate shared feature mapping');
      selections.set(entry.path,entry.scenarioIds);
    }
    if(Object.hasOwn(shared,'catalogue')){
      const cataloguePath=sharedPrefix+'ledgers/workflows.json';
      if(shared.catalogue!==cataloguePath)throw Error('Invalid shared catalogue path');
      const text=await Bun.file(join(referenceRoot,'ledgers/workflows.json')).text(),ledger=JSON.parse(text);
      if(ledger.schemaVersion!==1||!Array.isArray(ledger.features)||!ledger.features.length||!Array.isArray(ledger.workflows)||!ledger.workflows.length)throw Error('Invalid shared catalogue inventory');
      if(ledger.features.some((p:unknown)=>!safeFeature(p))||new Set(ledger.features).size!==ledger.features.length)throw Error('Invalid or duplicate catalogue feature path');
      const discovered=await Array.fromAsync(new Bun.Glob('workflows/**/*.feature').scan({cwd:referenceRoot,onlyFiles:true}));
      const manifest=await Bun.file(join(referenceRoot,'manifest.json')).json();
      if(!Array.isArray(manifest.files))throw Error('Invalid catalogue manifest');
      const sealed=manifest.files.filter((a:any)=>a.role==='workflow').map((a:any)=>a.path);
      const exact=(paths:string[])=>paths.length===ledger.features.length&&new Set(paths).size===paths.length&&paths.every(p=>ledger.features.includes(p));
      if(!exact(discovered)||!exact(sealed))throw Error('Shared catalogue feature membership or seal drift');
      const owners=new Map<string,{feature:string;expandedCases:number}>();
      for(const row of ledger.workflows){
        if(!row||typeof row.id!=='string'||!row.id.startsWith('@id-')||owners.has(row.id)||!ledger.features.includes(row.feature)||!Number.isSafeInteger(row.expandedCases)||row.expandedCases<1)throw Error('Invalid or duplicate catalogue scenario');
        owners.set(row.id,row);
      }
      for(const path of selections.keys())if(!ledger.features.includes(path.slice(sharedPrefix.length)))throw Error('Activation outside shared catalogue');
      if(pptxCandidate){for(const id of [...pptxManipulationPin.selectedScenarioIds,...(formattingCandidate?pptxFormattingPin.selectedScenarioIds:[])]){
        const owner=owners.get(id);if(!owner)throw Error('Missing PPTX candidate scenario: '+id);
        const uri=sharedPrefix+owner.feature,ids=selections.get(uri) as string[]|undefined;
        selections.set(uri,[...(ids??[]),id]);
      }}
      const found=new Set<string>();let caseCount=0;
      for(const path of ledger.features as string[]){
        const uri=sharedPrefix+path,source=await Bun.file(join(referenceRoot,path)).text();
        const parsed=selections.has(uri)?selectSharedScenarios(uri,source,selections.get(uri)):parseFeature(uri,source,{allowDuplicateCaseNames:true});
        // Unselected files must remain planned in the canonical source.
        if(!selections.has(uri)&&parsed.lifecycle!=='planned')throw Error('Shared feature must remain planned: '+path);
        for(const scenario of parsed.scenarios){const owner=owners.get(scenario.scenarioId);
          if(found.has(scenario.scenarioId)||!owner||owner.feature!==path||owner.expandedCases!==scenario.cases.length)throw Error('Shared catalogue scenario ownership or case-count drift: '+scenario.scenarioId);
          found.add(scenario.scenarioId);caseCount+=scenario.cases.length;
        }
        features.push(parsed);
      }
      if(found.size!==owners.size)throw Error('Shared catalogue scenario missing from source');
      catalogue={path:cataloguePath,sourceSha256:sha256(text),features:ledger.features.length,scenarios:found.size,cases:caseCount};
    }else{
      // Small isolated test projects may intentionally select only a few files.
      for(const [path,ids]of selections){const text=await Bun.file(join(referenceRoot,path.slice(sharedPrefix.length))).text();features.push(selectSharedScenarios(path,text,ids));}
    }
  }
  features.sort((left, right) => left.path.localeCompare(right.path));
  validateInventory(features);
  return {
    root,
    features,
    counts: countInventory(features),
    coverage: {catalogue,shared:countInventory(features.filter(f=>f.path.startsWith(sharedPrefix))),localOnly:countInventory(features.filter(f=>!f.path.startsWith(sharedPrefix)))},
  };
}

export function matchStepBinding(
  bindings: readonly StepBinding[],
  text: string,
): BindingMatch[] {
  return bindings.flatMap((binding) => {
    const captures = exactMatch(binding.pattern, text);
    return captures ? [{ binding, captures }] : [];
  });
}

export async function executeAcceptance(
  inventory: AcceptanceInventory,
  bindings: readonly StepBinding[],
  runId: string,
): Promise<AcceptanceExecution> {
  const features: AcceptanceFeatureReport[] = [];
  const failures: string[] = [];

  for (const feature of inventory.features) {
    if (feature.lifecycle === "planned") {
      features.push(planFeature(feature));
      continue;
    }

    const scenarioReports: AcceptanceScenarioReport[] = [];
    for (const scenario of feature.scenarios) {
      const lifecycle = scenario.lifecycle ?? feature.lifecycle;
      if (lifecycle === 'planned') {
        scenarioReports.push(planFeature({...feature, lifecycle: 'planned', scenarios: [scenario]}).scenarios[0]!);
        continue;
      }
      const caseReports: AcceptanceCaseReport[] = [];
      for (const acceptanceCase of scenario.cases) {
        const state: Record<string, unknown> = {};
        const context: Record<string, unknown> = {
          runId,
          root: inventory.root,
          state,
          feature: {
            path: feature.path,
            name: feature.name,
            lifecycle: feature.lifecycle,
            runner: feature.runner,
            tags: [...feature.tags],
            sourceSha256: feature.sourceSha256,
          },
          scenario: {
            id: scenario.scenarioId,
            lifecycle,
            name: scenario.name,
            tags: [...scenario.tags],
            location: scenario.location,
          },
          case: {
            id: acceptanceCase.caseId,
            lifecycle,
            name: acceptanceCase.name,
            location: acceptanceCase.location,
            tags: [...acceptanceCase.tags],
            example: acceptanceCase.example
              ? {
                  values: { ...acceptanceCase.example.values },
                  orderedPairs: acceptanceCase.example.orderedPairs.map((pair) => ({ ...pair })),
                  location: acceptanceCase.example.location,
                }
              : undefined,
          },
        };

        let caseError: string | undefined;
        let stopIndex = acceptanceCase.steps.length;
        const steps: AcceptanceStepReport[] = [];

        for (const [index, step] of acceptanceCase.steps.entries()) {
          context.step = {
            index,
            keyword: step.keyword,
            text: step.text,
            type: step.type,
            source: step.source,
            fromBackground: step.fromBackground,
            argument: step.argument,
          };

          const matches = matchStepBinding(bindings, step.text);
          if (!matches.length) {
            const error = `Undefined step: ${step.text}`;
            steps.push({ ...step, status: "undefined", error });
            caseError = error;
            stopIndex = index + 1;
            break;
          }
          if (matches.length > 1) {
            const error = `Ambiguous step: ${step.text}`;
            steps.push({
              ...step,
              status: "ambiguous",
              error,
              binding: matches.map((match) => String(match.binding.pattern)).join(", "),
            });
            caseError = error;
            stopIndex = index + 1;
            break;
          }

          const match = matches[0];
          if (!match) {
            throw new Error(`Missing binding match after successful lookup: ${step.text}`);
          }
          const { binding, captures } = match;
          const startedAt = Date.now();
          try {
            await binding.run(context, ...captures);
            steps.push({
              ...step,
              status: "passed",
              binding: String(binding.pattern),
              captures,
              durationMs: Math.max(0, Date.now() - startedAt),
            });
          } catch (error) {
            const message = formatError(error);
            steps.push({
              ...step,
              status: "failed",
              binding: String(binding.pattern),
              captures,
              durationMs: Math.max(0, Date.now() - startedAt),
              error: message,
            });
            caseError = message;
            stopIndex = index + 1;
            break;
          }
        }

        for (const step of acceptanceCase.steps.slice(stopIndex)) {
          steps.push({
            ...step,
            status: "skipped",
            error: caseError ? `Skipped after failure: ${caseError}` : undefined,
          });
        }

        const result: AcceptanceResult = caseError ? "failed" : "passed";
        if (caseError) {
          failures.push(
            `${feature.path} :: ${scenario.scenarioId} :: ${acceptanceCase.caseId} :: ${caseError}`,
          );
        }
        caseReports.push({
          ...acceptanceCase,
          lifecycle,
          result,
          steps,
          error: caseError,
        });
      }

      scenarioReports.push({
        scenarioId: scenario.scenarioId,
        name: scenario.name,
        location: scenario.location,
        tags: [...scenario.tags],
        lifecycle,
        result: caseReports.every((acceptanceCase) => acceptanceCase.result === "passed")
          ? "passed"
          : "failed",
        cases: caseReports,
      });
    }

    features.push({
      path: feature.path,
      name: feature.name,
      lifecycle: feature.lifecycle,
      runner: feature.runner,
      tags: [...feature.tags],
      sourceSha256: feature.sourceSha256,
      // Feature result summarises attempted scenarios; planned cases stay explicit below.
      result: scenarioReports.some(scenario => scenario.result === 'failed') ? 'failed' : 'passed',
      scenarios: scenarioReports,
    });
  }

  return {
    features,
    counts: countExecution(features),
    failures,
  };
}

function planFeature(feature: AcceptanceFeature): AcceptanceFeatureReport {
  return {
    path: feature.path,
    name: feature.name,
    lifecycle: feature.lifecycle,
    runner: feature.runner,
    tags: [...feature.tags],
    sourceSha256: feature.sourceSha256,
    result: "planned",
    scenarios: feature.scenarios.map((scenario) => ({
      scenarioId: scenario.scenarioId,
      name: scenario.name,
      location: scenario.location,
      tags: [...scenario.tags],
      lifecycle: feature.lifecycle,
      result: "planned",
      cases: scenario.cases.map((acceptanceCase) => ({
        pickleId: acceptanceCase.pickleId,
        caseId: acceptanceCase.caseId,
        identityKey: acceptanceCase.identityKey,
        name: acceptanceCase.name,
        location: acceptanceCase.location,
        example: acceptanceCase.example
          ? {
              values: { ...acceptanceCase.example.values },
              orderedPairs: acceptanceCase.example.orderedPairs.map((pair) => ({ ...pair })),
              location: acceptanceCase.example.location,
            }
          : undefined,
        tags: [...acceptanceCase.tags],
        lifecycle: feature.lifecycle,
        result: "planned",
        steps: acceptanceCase.steps.map((step) => ({
          ...step,
          status: "planned",
        })),
      })),
    })),
  };
}

function buildCase(
  path: string,
  sourceSha256: string,
  index: DocumentIndex,
  scenarioId: string,
  scenario: Scenario,
  pickle: Pickle,
): AcceptanceCase {
  const rowId = pickle.astNodeIds.find((id) => id !== scenario.id && index.rows.has(id));
  const row = rowId ? index.rows.get(rowId) : undefined;
  const orderedPairs = row?.header.map((name, offset) => ({
    name,
    value: row.values[offset] ?? "",
  }));
  const valuesRecord = row ? { ...row.valuesRecord } : undefined;
  const caseId = row
    ? `${scenarioId} -- ${orderedPairs
        ?.map((pair) => `${pair.name}=${JSON.stringify(pair.value)}`)
        .join(", ")}`
    : scenarioId;
  const identityKey = row ? `${scenarioId}|${JSON.stringify(row.valuesRecord)}` : scenarioId;

  return {
    pickleId: pickle.id,
    caseId,
    identityKey,
    name: pickle.name.trim(),
    location: toSourceLocation(path, pickle.location ?? scenario.location),
    example: row && orderedPairs
      ? {
          values: valuesRecord ?? {},
          orderedPairs,
          location: toSourceLocation(path, row.location),
        }
      : undefined,
    tags: pickle.tags.map((tag) => tag.name),
    steps: pickle.steps.map((pickleStep) =>
      buildStep(path, sourceSha256, index, row, pickleStep),
    ),
  } satisfies AcceptanceCase;
}

function buildStep(
  path: string,
  sourceSha256: string,
  index: DocumentIndex,
  row: IndexedRow | undefined,
  pickleStep: PickleStep,
): AcceptanceStep {
  const stepId = pickleStep.astNodeIds.find((id) => index.steps.has(id));
  if (!stepId) {
    throw new Error(`Missing source step for compiled pickle step in ${path}`);
  }
  const indexed = index.steps.get(stepId);
  if (!indexed) {
    throw new Error(`Missing step index for ${stepId} in ${path}`);
  }
  const text = pickleStep.text.trim();
  if (!text) {
    throw new Error(
      `Step text must not be empty: ${path}:${indexed.step.location.line}`,
    );
  }
  return {
    keyword: indexed.step.keyword.trim(),
    text,
    type: pickleStep.type ?? PickleStepType.UNKNOWN,
    source: {
      sha256: sourceSha256,
      location: toSourceLocation(path, indexed.step.location),
      exampleLocation: row ? toSourceLocation(path, row.location) : undefined,
    },
    fromBackground: indexed.fromBackground,
    argument: pickleStep.argument
      ? toStepArgumentView(pickleStep.argument)
      : undefined,
  };
}

function indexDocument(path: string, feature: Feature): DocumentIndex {
  const steps = new Map<string, IndexedStep>();
  const rows = new Map<string, IndexedRow>();

  const indexScenario = (scenario: Scenario) => {
    for (const step of scenario.steps) {
      steps.set(step.id, { step, fromBackground: false });
    }
    for (const examples of scenario.examples) {
      const headerCells = examples.tableHeader?.cells.map((cell) => cell.value.trim()) ?? [];
      if (examples.tableHeader && !headerCells.length) {
        throw new Error(`Examples table header is empty: ${path}:${examples.location.line}`);
      }
      if (new Set(headerCells).size !== headerCells.length) {
        throw new Error(`Examples table header contains duplicates: ${path}:${examples.location.line}`);
      }
      if (headerCells.some((cell) => !cell)) {
        throw new Error(`Examples table header contains empty cells: ${path}:${examples.location.line}`);
      }
      for (const row of examples.tableBody) {
        rows.set(row.id, {
          header: headerCells,
          values: row.cells.map((cell) => cell.value),
          valuesRecord: Object.fromEntries(
            headerCells.map((name, offset) => [name, row.cells[offset]?.value ?? ""]),
          ),
          location: row.location,
        });
      }
    }
  };

  const indexBackground = (stepsToIndex: readonly Step[]) => {
    for (const step of stepsToIndex) {
      steps.set(step.id, { step, fromBackground: true });
    }
  };

  const indexRule = (rule: Rule) => {
    for (const child of rule.children) {
      if (child.background) {
        indexBackground(child.background.steps);
      }
      if (child.scenario) {
        indexScenario(child.scenario);
      }
    }
  };

  for (const child of feature.children) {
    if (child.background) {
      indexBackground(child.background.steps);
    }
    if (child.rule) {
      indexRule(child.rule);
    }
    if (child.scenario) {
      indexScenario(child.scenario);
    }
  }

  return { steps, rows };
}

function collectScenarios(feature: Feature): Scenario[] {
  const scenarios: Scenario[] = [];
  for (const child of feature.children) {
    if (child.scenario) {
      scenarios.push(child.scenario);
    }
    if (child.rule) {
      for (const ruleChild of child.rule.children) {
        if (ruleChild.scenario) {
          scenarios.push(ruleChild.scenario);
        }
      }
    }
  }
  return scenarios;
}

function validateFeatureTags(
  path: string,
  tags: readonly string[],
): AcceptanceLifecycle {
  const unique = [...new Set(tags)];
  if (
    tags.length === 2 &&
    unique.length === 2 &&
    unique.includes("@implemented") &&
    unique.includes("@bun")
  ) {
    return "implemented";
  }
  if (tags.length === 1 && unique.length === 1 && unique[0] === "@planned") {
    return "planned";
  }
  throw new Error(
    `Feature tags must be exactly @implemented @bun or @planned: ${path}`,
  );
}

function validateScenarioTags(
  path: string,
  scenario: Scenario,
  tags: readonly string[],
): string {
  if (tags.some((tag) => RESERVED_SCENARIO_TAGS.has(tag))) {
    throw new Error(
      `Scenario tags cannot override lifecycle or runner: ${path}:${scenario.location.line}`,
    );
  }
  const idTags = tags.filter((tag) => tag.startsWith("@id-"));
  if (idTags.length !== 1) {
    throw new Error(
      `Scenario must declare exactly one @id-* tag: ${path}:${scenario.location.line}`,
    );
  }
  return idTags[0] ?? "";
}

function validateExamplesTags(path: string, examples: Examples): void {
  if (examples.tags.some((tag) => RESERVED_SCENARIO_TAGS.has(tag.name) || tag.name.startsWith("@id-"))) {
    throw new Error(
      `Examples tags cannot override lifecycle, runner or scenario id: ${path}:${examples.location.line}`,
    );
  }
}

function validateInventory(features: readonly AcceptanceFeature[]): void {
  const featureNames = new Map<string, string>();
  const scenarioIds = new Map<string, string>();
  const caseIds = new Map<string, string>();

  for (const feature of features) {
    const priorFeature = featureNames.get(feature.name);
    if (priorFeature) {
      throw new Error(
        `Duplicate feature name ${JSON.stringify(feature.name)} in ${priorFeature} and ${feature.path}`,
      );
    }
    featureNames.set(feature.name, feature.path);

    for (const scenario of feature.scenarios) {
      const priorScenario = scenarioIds.get(scenario.scenarioId);
      if (priorScenario) {
        throw new Error(
          `Duplicate scenario id ${scenario.scenarioId} in ${priorScenario} and ${feature.path}`,
        );
      }
      scenarioIds.set(scenario.scenarioId, feature.path);

      for (const acceptanceCase of scenario.cases) {
        const priorCase = caseIds.get(acceptanceCase.identityKey);
        if (priorCase) {
          throw new Error(
            `Duplicate case identity ${acceptanceCase.caseId} in ${priorCase} and ${feature.path}`,
          );
        }
        caseIds.set(acceptanceCase.identityKey, feature.path);
      }
    }
  }
}

function countInventory(features: readonly AcceptanceFeature[]) {
  const counts = {
    features: zeroInventoryCount(),
    scenarios: zeroInventoryCount(),
    cases: zeroInventoryCount(),
    steps: zeroInventoryCount(),
  };

  for (const feature of features) {
    bumpInventoryCount(counts.features, feature.lifecycle, 1);
    for (const scenario of feature.scenarios) {
      const lifecycle = scenario.lifecycle ?? feature.lifecycle;
      bumpInventoryCount(counts.scenarios, lifecycle, 1);
      for (const acceptanceCase of scenario.cases) {
        bumpInventoryCount(counts.cases, lifecycle, 1);
        bumpInventoryCount(counts.steps, lifecycle, acceptanceCase.steps.length);
      }
    }
  }

  return counts;
}

function countExecution(features: readonly AcceptanceFeatureReport[]) {
  const counts = {
    features: zeroExecutionCount(),
    scenarios: zeroExecutionCount(),
    cases: zeroExecutionCount(),
    steps: zeroStepExecutionCount(),
  };

  for (const feature of features) {
    bumpExecutionCount(counts.features, feature.result);
    for (const scenario of feature.scenarios) {
      bumpExecutionCount(counts.scenarios, scenario.result);
      for (const acceptanceCase of scenario.cases) {
        bumpExecutionCount(counts.cases, acceptanceCase.result);
        for (const step of acceptanceCase.steps) {
          bumpStepExecutionCount(counts.steps, step.status);
        }
      }
    }
  }

  return counts;
}

function exactMatch(pattern: RegExp, text: string): string[] | undefined {
  const regex = new RegExp(pattern.source, pattern.flags.replace(/[gy]/g, ""));
  const match = regex.exec(text);
  if (!match || match[0] !== text) {
    return undefined;
  }
  return match.slice(1).map((capture) => capture ?? "");
}

function toStepArgumentView(argument: PickleStepArgument): StepArgumentView {
  return {
    docString: argument.docString
      ? {
          content: argument.docString.content,
          mediaType: argument.docString.mediaType,
        }
      : undefined,
    dataTable: argument.dataTable
      ? argument.dataTable.rows.map((row) => row.cells.map((cell) => cell.value))
      : undefined,
  };
}

function toSourceLocation(uri: string, location: Location): SourceLocation {
  return {
    uri,
    line: location.line,
    column: location.column,
  };
}

function formatError(error: unknown): string {
  if (error instanceof Error) {
    return error.stack ?? error.message;
  }
  return String(error);
}

function sha256(text: string): string {
  return createHash("sha256").update(text).digest("hex");
}

function zeroInventoryCount(): InventoryCount {
  return { implemented: 0, planned: 0, total: 0 };
}

function zeroExecutionCount(): ExecutionCount {
  return { passed: 0, failed: 0, planned: 0, total: 0 };
}

function zeroStepExecutionCount(): StepExecutionCount {
  return {
    passed: 0,
    failed: 0,
    planned: 0,
    skipped: 0,
    undefined: 0,
    ambiguous: 0,
    total: 0,
  };
}

function bumpInventoryCount(
  count: InventoryCount,
  lifecycle: AcceptanceLifecycle,
  amount: number,
): void {
  count[lifecycle] += amount;
  count.total += amount;
}

function bumpExecutionCount(count: ExecutionCount, result: AcceptanceResult): void {
  count[result] += 1;
  count.total += 1;
}

function bumpStepExecutionCount(
  count: StepExecutionCount,
  result: AcceptanceStepResult,
): void {
  count[result] += 1;
  count.total += 1;
}
