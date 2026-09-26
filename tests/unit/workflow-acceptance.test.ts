import { afterAll, describe, expect, test } from "bun:test";
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import { PickleStepType } from "@cucumber/messages";

import { executeAcceptance, newAcceptanceRunId, parseFeature, type AcceptanceFeature, type AcceptanceInventory } from "../../scripts/gherkin.ts";
import { bindings, cleanupWorkflowFixtures } from "../acceptance/workflow.ts";

const PROJECT_ROOT = new URL("../..", import.meta.url).pathname;
const FEATURE_PATH = "features/native/workflow-mutation-safety.feature";
const FROZEN_FEATURE = join(fixturesRoot(),'shared/v2/pack/features/mutation-safety.feature');

afterAll(async () => {
  await cleanupWorkflowFixtures();
});

describe("workflow shared acceptance", () => {
  test("executes the frozen shared workflow feature as 19 implemented Bun cases", async () => {
    const text = (await Bun.file(FROZEN_FEATURE).text()).replace(/^@planned\s*$/m, "@implemented @bun");
    const feature = parseFeature(FEATURE_PATH, text);
    const inventory = inventoryFor(feature);
    const caseCount = feature.scenarios.flatMap((scenario) => scenario.cases).length;
    const expectedOutcomeCount = feature.scenarios
      .flatMap((scenario) => scenario.cases)
      .flatMap((acceptanceCase) => acceptanceCase.steps)
      .filter((step) => step.type === PickleStepType.OUTCOME).length;

    expect(feature.lifecycle).toBe("implemented");
    expect(feature.runner).toBe("bun");
    expect(feature.path).toBe(FEATURE_PATH);
    expect(feature.path).not.toContain("/planned/");
    expect(feature.scenarios).toHaveLength(8);
    expect(caseCount).toBe(19);

    const execution = await executeAcceptance(inventory, bindings, newAcceptanceRunId());
    const passedOutcomeCount = execution.features
      .flatMap((report) => report.scenarios)
      .flatMap((report) => report.cases)
      .flatMap((report) => report.steps)
      .filter((step) => step.type === PickleStepType.OUTCOME && step.status === "passed").length;

    expect(execution.failures).toEqual([]);
    expect(execution.counts.features).toEqual({ passed: 1, failed: 0, planned: 0, total: 1 });
    expect(execution.counts.scenarios).toEqual({ passed: 8, failed: 0, planned: 0, total: 8 });
    expect(execution.counts.cases.passed).toBe(19);
    expect(execution.counts.cases.failed).toBe(0);
    expect(execution.counts.cases.planned).toBe(0);
    expect(execution.counts.cases.total).toBe(19);
    expect(passedOutcomeCount).toBe(expectedOutcomeCount);
    expect(execution.counts.steps.failed).toBe(0);
    expect(execution.counts.steps.undefined).toBe(0);
    expect(execution.counts.steps.ambiguous).toBe(0);
  }, 300_000);
});

function inventoryFor(feature: AcceptanceFeature): AcceptanceInventory {
  const scenarios = feature.scenarios.length;
  const cases = feature.scenarios.flatMap((scenario) => scenario.cases).length;
  const steps = feature.scenarios
    .flatMap((scenario) => scenario.cases)
    .flatMap((acceptanceCase) => acceptanceCase.steps).length;
  return {
    root: PROJECT_ROOT,
    features: [feature],
    counts: {
      features: { implemented: 1, planned: 0, total: 1 },
      scenarios: { implemented: scenarios, planned: 0, total: scenarios },
      cases: { implemented: cases, planned: 0, total: cases },
      steps: { implemented: steps, planned: 0, total: steps },
    },
  };
}
