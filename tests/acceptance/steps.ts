import type { StepBinding } from "../../scripts/gherkin.ts";
import { bindings as coreBindings } from "./core.ts";
import { bindings as zip64Bindings } from "./zip64.ts";
import { bindings as namespaceBindings } from "./relationship-namespaces.ts";
import { bindings as docxBindings } from "./docx.ts";
import { bindings as pptxBindings } from "./pptx.ts";
import { bindings as xlsxBindings } from "./xlsx.ts";
import { bindings as workflowBindings, cleanupWorkflowFixtures } from "./workflow.ts";
export const cleanup = cleanupWorkflowFixtures;

export const bindings: StepBinding[] = [
  ...coreBindings,
  ...zip64Bindings,
  ...namespaceBindings,
  ...docxBindings,
  ...pptxBindings,
  ...xlsxBindings,
  ...workflowBindings,
];

export default bindings;
