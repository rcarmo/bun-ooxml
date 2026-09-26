import type { StepBinding } from "../../scripts/gherkin.ts";
import { bindings as coreBindings } from "./core.ts";
import { bindings as xmlNameBindings } from "./xml-names.ts";
import { bindings as graphBindings } from "./graph.ts";
import { bindings as zip64Bindings } from "./zip64.ts";
import { bindings as namespaceBindings } from "./relationship-namespaces.ts";
import { bindings as docxBindings } from "./docx.ts";
import { bindings as storyBindings } from "./story-docx.ts";
import { bindings as revisionBindings } from "./revisions-docx.ts";
import { bindings as redlineBindings } from "./redline-docx.ts";
import { bindings as commentBindings } from "./comments-docx.ts";
import { bindings as formattingBindings } from './run-formatting.ts';
import { bindings as trackedWorkflowBindings, cleanupTrackedWorkflowFixtures } from './tracked-workflow.ts';
import { bindings as createDocxBindings } from "./create-docx.ts";
import { bindings as createPptxBindings } from "./create-pptx.ts";
import { bindings as createXlsxBindings } from "./create-xlsx.ts";
import { bindings as tableDocxBindings } from "./tables-docx.ts";
import { bindings as tablePptxBindings } from "./tables-pptx.ts";
import { bindings as pptxBindings } from "./pptx.ts";
import { bindings as xlsxBindings } from "./xlsx.ts";
import { bindings as cacheBoundaryBindings } from "./cache-boundaries.ts";
import { bindings as workflowBindings, cleanupWorkflowFixtures } from "./workflow.ts";
export async function cleanup(){await Promise.all([cleanupWorkflowFixtures(),cleanupTrackedWorkflowFixtures()]);}

export const bindings: StepBinding[] = [
  ...coreBindings,
  ...xmlNameBindings,
  ...graphBindings,
  ...zip64Bindings,
  ...namespaceBindings,
  ...docxBindings,
  ...storyBindings,
  ...revisionBindings,
  ...redlineBindings,
  ...commentBindings,
  ...formattingBindings,
  ...trackedWorkflowBindings,
  ...createDocxBindings,
  ...createPptxBindings,
  ...createXlsxBindings,
  ...tableDocxBindings,
  ...tablePptxBindings,
  ...pptxBindings,
  ...xlsxBindings,
  ...cacheBoundaryBindings,
  ...workflowBindings,
];

export default bindings;
