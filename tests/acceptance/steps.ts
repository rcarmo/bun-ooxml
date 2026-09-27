import type { StepBinding } from "../../scripts/gherkin.ts";
import { bindings as coreBindings } from "./core.ts";
import { bindings as packageAdmissionBindings } from './package-admission.ts';
import { bindings as packageComparisonBindings } from './package-comparison.ts';
import { bindings as refusalBindings } from './refusal-outcomes.ts';
import { bindings as opcCustodyBindings, cleanup as cleanupOpcCustody } from './opc-custody.ts';
import { bindings as xmlNameBindings } from "./xml-names.ts";
import { bindings as xmlComparisonBindings } from './xml-comparison.ts';
import { bindings as xmlRemovalBindings } from './xml-removal.ts';
import { bindings as xmlValuesBindings } from './xml-values.ts';
import { bindings as xmlAttributeBindings } from './xml-attributes.ts';
import { bindings as xmlStructureBindings } from './xml-structure.ts';
import { bindings as graphBindings } from "./graph.ts";
import { bindings as zip64Bindings } from "./zip64.ts";
import { bindings as zip32Bindings } from './zip32.ts';
import { bindings as namespaceBindings } from "./relationship-namespaces.ts";
import { bindings as docxBindings } from "./docx.ts";
import { bindings as storyBindings } from "./story-docx.ts";
import { bindings as revisionBindings } from "./revisions-docx.ts";
import { bindings as redlineBindings } from "./redline-docx.ts";
import { bindings as commentBindings } from "./comments-docx.ts";
import { bindings as formattingBindings } from './run-formatting.ts';
import { bindings as fontSizeBindings, cleanup as cleanupFontSize } from './font-size.ts';
import { bindings as paragraphStyleBindings } from './paragraph-style.ts';
import { bindings as paragraphPropertiesBindings } from './paragraph-properties.ts';
import { bindings as styleAuthoringBindings } from './style-authoring.ts';
import { bindings as textBoxBindings } from './text-box.ts';
import { bindings as cellStyleBindings } from './cell-style.ts';
import { bindings as pageLayoutBindings } from './page-layout.ts';
import { bindings as slideOrderBindings } from './slide-order.ts';
import { bindings as effectiveFormattingBindings } from './effective-formatting.ts';
import { bindings as trackedWorkflowBindings, cleanupTrackedWorkflowFixtures } from './tracked-workflow.ts';
import { bindings as createDocxBindings } from "./create-docx.ts";
import { bindings as createPptxBindings } from "./create-pptx.ts";
import { bindings as createXlsxBindings } from "./create-xlsx.ts";
import { bindings as tableDocxBindings } from "./tables-docx.ts";
import { bindings as tablePptxBindings } from "./tables-pptx.ts";
import { bindings as pptxBindings } from "./pptx.ts";
import { bindings as pptxCustodyBindings } from './pptx-custody.ts';
import { bindings as notesEditingBindings } from './notes-editing.ts';
import { bindings as xlsxBindings } from "./xlsx.ts";
import { bindings as xlsxRangeBindings } from './xlsx-range.ts';
import { bindings as cacheBoundaryBindings } from "./cache-boundaries.ts";
import { bindings as workflowBindings, cleanupWorkflowFixtures } from "./workflow.ts";
export async function cleanup(){await Promise.all([cleanupWorkflowFixtures(),cleanupTrackedWorkflowFixtures(),cleanupFontSize(),cleanupOpcCustody()]);}

export const bindings: StepBinding[] = [
  ...coreBindings,
  ...packageAdmissionBindings,
  ...packageComparisonBindings,
  ...refusalBindings,
  ...opcCustodyBindings,
  ...xmlNameBindings,
  ...xmlComparisonBindings,
  ...xmlRemovalBindings,
  ...xmlValuesBindings,
  ...xmlAttributeBindings,
  ...xmlStructureBindings,
  ...graphBindings,
  ...zip64Bindings,
  ...zip32Bindings,
  ...namespaceBindings,
  ...docxBindings,
  ...storyBindings,
  ...revisionBindings,
  ...redlineBindings,
  ...commentBindings,
  ...formattingBindings,
  ...fontSizeBindings,
  ...paragraphStyleBindings,
  ...paragraphPropertiesBindings,
  ...styleAuthoringBindings,
  ...textBoxBindings,
  ...cellStyleBindings,
  ...pageLayoutBindings,
  ...slideOrderBindings,
  ...effectiveFormattingBindings,
  ...trackedWorkflowBindings,
  ...createDocxBindings,
  ...createPptxBindings,
  ...createXlsxBindings,
  ...tableDocxBindings,
  ...tablePptxBindings,
  ...pptxBindings,
  ...pptxCustodyBindings,
  ...notesEditingBindings,
  ...xlsxBindings,
  ...xlsxRangeBindings,
  ...cacheBoundaryBindings,
  ...workflowBindings,
];

export default bindings;
