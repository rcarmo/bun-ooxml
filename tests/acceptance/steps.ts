import type { StepBinding } from "../../scripts/gherkin.ts";
import { bindings as coreBindings } from "./core.ts";
import { bindings as graphBindings } from "./graph.ts";
import { bindings as zip64Bindings } from "./zip64.ts";
import { bindings as namespaceBindings } from "./relationship-namespaces.ts";
import { bindings as docxBindings } from "./docx.ts";
import { bindings as createDocxBindings } from "./create-docx.ts";
import { bindings as createPptxBindings } from "./create-pptx.ts";
import { bindings as createXlsxBindings } from "./create-xlsx.ts";
import { bindings as pptxBindings } from "./pptx.ts";
import { bindings as xlsxBindings } from "./xlsx.ts";
import { bindings as cacheBoundaryBindings } from "./cache-boundaries.ts";
import { bindings as workflowBindings, cleanupWorkflowFixtures } from "./workflow.ts";
export const cleanup = cleanupWorkflowFixtures;

export const bindings: StepBinding[] = [
  ...coreBindings,
  ...graphBindings,
  ...zip64Bindings,
  ...namespaceBindings,
  ...docxBindings,
  ...createDocxBindings,
  ...createPptxBindings,
  ...createXlsxBindings,
  ...pptxBindings,
  ...xlsxBindings,
  ...cacheBoundaryBindings,
  ...workflowBindings,
];

export default bindings;
