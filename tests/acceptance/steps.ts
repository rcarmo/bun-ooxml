import type { StepBinding } from "../../scripts/gherkin.ts";
import { bindings as coreBindings } from "./core.ts";
import { bindings as docxBindings } from "./docx.ts";
import { bindings as pptxBindings } from "./pptx.ts";
import { bindings as xlsxBindings } from "./xlsx.ts";
import { bindings as workflowBindings, cleanupWorkflowFixtures } from "./workflow.ts";
export const cleanup = cleanupWorkflowFixtures;

export const bindings: StepBinding[] = [
  ...coreBindings,
  ...docxBindings,
  ...pptxBindings,
  ...xlsxBindings,
  ...workflowBindings,
];

export default bindings;
