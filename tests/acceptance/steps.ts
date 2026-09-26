import type { StepBinding } from "../../scripts/gherkin.ts";
import { bindings as coreBindings } from "./core.ts";
import { bindings as docxBindings } from "./docx.ts";
import { bindings as pptxBindings } from "./pptx.ts";
import { bindings as xlsxBindings } from "./xlsx.ts";

export const bindings: StepBinding[] = [
  ...coreBindings,
  ...docxBindings,
  ...pptxBindings,
  ...xlsxBindings,
];

export default bindings;
