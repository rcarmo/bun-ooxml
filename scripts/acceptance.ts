import { mkdir } from "node:fs/promises";
import { dirname, join } from "node:path";
import { pathToFileURL } from "node:url";
import {
  executeAcceptance,
  inventoryFeatures,
  newAcceptanceRunId,
  type AcceptanceExecution,
  type AcceptanceFeatureReport,
  type AcceptanceInventory,
  type InventoryCount,
  type StepBinding,
} from "./gherkin.ts";

export type AcceptanceRunReport = {
  runId: string;
  root: string;
  artifactPath: string;
  stepsModulePath: string;
  full: boolean;
  startedAt: string;
  finishedAt?: string;
  status: "running" | "passed" | "failed";
  inventory: AcceptanceInventory["counts"];
  execution: AcceptanceExecution["counts"];
  features: AcceptanceFeatureReport[];
  failures: string[];
};

export async function runAcceptance(
  bindings?: readonly StepBinding[],
  options: {
    root?: string;
    full?: boolean;
    artifactPath?: string;
    stepsModulePath?: string;
  } = {},
): Promise<AcceptanceRunReport> {
  const root = options.root ?? join(import.meta.dir, "..");
  const artifactPath = options.artifactPath ?? join(root, "artifacts", "acceptance.json");
  const stepsModulePath = options.stepsModulePath ?? join(root, "tests", "acceptance", "steps.ts");
  const runId = newAcceptanceRunId();
  const full = options.full ?? false;

  const report: AcceptanceRunReport = {
    runId,
    root,
    artifactPath,
    stepsModulePath,
    full,
    startedAt: new Date().toISOString(),
    status: "running",
    inventory: zeroInventoryCounts(),
    execution: zeroExecutionCounts(),
    features: [],
    failures: [],
  };

  await mkdir(dirname(artifactPath), { recursive: true });
  await writeReport(report);

  let cleanup: (() => unknown | Promise<unknown>) | undefined;
  try {
    const inventory = await inventoryFeatures(root);
    report.inventory = inventory.counts;

    const loaded = !bindings && inventory.counts.cases.implemented > 0
      ? await loadBindings(stepsModulePath, runId) : undefined;
    cleanup=loaded?.cleanup;
    const effectiveBindings=bindings??loaded?.bindings??[];

    const execution = await executeAcceptance(inventory, effectiveBindings, runId);
    report.execution = execution.counts;
    report.features = execution.features;
    report.failures.push(...execution.failures);

    if (report.inventory.cases.implemented === 0) {
      report.failures.push("No implemented cases executed; inventory alone cannot pass acceptance.");
    }

    if (full && report.inventory.cases.planned > 0) {
      report.failures.push(
        `Full acceptance rejects planned features or cases: ${report.inventory.cases.planned} planned case(s) remain.`,
      );
    }

    report.status = report.failures.length ? "failed" : "passed";
  } catch (error) {
    report.status = "failed";
    report.failures.push(formatError(error));
  }

  try { await cleanup?.(); } catch(error) {
    report.status='failed'; report.failures.push(`Acceptance cleanup failed: ${formatError(error)}`);
  }
  report.finishedAt = new Date().toISOString();
  await writeReport(report);

  if (report.status === "failed") {
    throw new Error(report.failures.join("\n\n"));
  }

  return report;
}

async function loadBindings(
  stepsModulePath: string,
  runId: string,
): Promise<{bindings:readonly StepBinding[]; cleanup?:()=>unknown|Promise<unknown>}> {
  try {
    const url = new URL(pathToFileURL(stepsModulePath).href);
    url.searchParams.set("acceptanceRunId", runId);
    const module = (await import(url.href)) as {
      bindings?: unknown;
      default?: unknown;
      cleanup?: unknown;
    };
    const candidate = module.bindings ?? module.default;
    if (!Array.isArray(candidate)) {
      throw new Error(`Expected ${stepsModulePath} to export a bindings array`);
    }
    const bindings = candidate as StepBinding[];
    for (const [index, binding] of bindings.entries()) {
      if (!(binding.pattern instanceof RegExp)) {
        throw new Error(
          `Binding ${index} in ${stepsModulePath} must expose pattern: RegExp`,
        );
      }
      if (typeof binding.run !== "function") {
        throw new Error(
          `Binding ${index} in ${stepsModulePath} must expose run(context, ...captures)`,
        );
      }
    }
    if(module.cleanup!==undefined&&typeof module.cleanup!=="function")throw new Error('cleanup export must be a function');
    return {bindings,cleanup:module.cleanup as (()=>unknown|Promise<unknown>)|undefined};
  } catch (error) {
    throw new Error(
      `Unable to load acceptance bindings from ${stepsModulePath}: ${formatError(error)}`,
    );
  }
}

async function writeReport(report: AcceptanceRunReport): Promise<void> {
  await Bun.write(report.artifactPath, `${JSON.stringify(report, null, 2)}\n`);
}

function zeroInventoryCounts(): {
  features: InventoryCount;
  scenarios: InventoryCount;
  cases: InventoryCount;
  steps: InventoryCount;
} {
  return {
    features: zeroInventoryCount(),
    scenarios: zeroInventoryCount(),
    cases: zeroInventoryCount(),
    steps: zeroInventoryCount(),
  };
}

function zeroInventoryCount(): InventoryCount {
  return { implemented: 0, planned: 0, total: 0 };
}

function zeroExecutionCounts(): AcceptanceExecution["counts"] {
  return {
    features: { passed: 0, failed: 0, planned: 0, total: 0 },
    scenarios: { passed: 0, failed: 0, planned: 0, total: 0 },
    cases: { passed: 0, failed: 0, planned: 0, total: 0 },
    steps: {
      passed: 0,
      failed: 0,
      planned: 0,
      skipped: 0,
      undefined: 0,
      ambiguous: 0,
      total: 0,
    },
  };
}

function formatError(error: unknown): string {
  if (error instanceof Error) {
    return error.stack ?? error.message;
  }
  return String(error);
}

if (import.meta.main) {
  const full = Bun.argv.slice(2).includes("--full");
  try {
    const report = await runAcceptance(undefined, { full });
    console.log(
      `Acceptance ${report.status}: ${report.execution.cases.passed}/${report.inventory.cases.implemented} implemented case(s) passed; planned cases ${report.inventory.cases.planned}.`,
    );
  } catch (error) {
    console.error(formatError(error));
    process.exitCode = 1;
  }
}
