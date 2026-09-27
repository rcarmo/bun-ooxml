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
import { bindings as dataDescriptorBindings } from './data-descriptor-integrity.ts';
import { bindings as namespaceBindings } from "./relationship-namespaces.ts";
import { bindings as docxBindings } from "./docx.ts";
import { bindings as storyBindings } from "./story-docx.ts";
import { bindings as revisionBindings } from "./revisions-docx.ts";
import { bindings as redlineBindings } from "./redline-docx.ts";
import { bindings as commentBindings } from "./comments-docx.ts";
import { bindings as formattingBindings } from './run-formatting.ts';
import { bindings as runEffectsBindings } from './run-effects.ts';
import { bindings as runAppearanceBindings } from './run-appearance.ts';
import { bindings as runFontNameBindings } from './run-font-name.ts';
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
import { bindings as trackingSettingsBindings } from './tracking-settings.ts';
import {bindings as trackingOutcomeBindings,cleanupTrackingOutcomes} from './tracking-outcomes.ts';
import { bindings as createDocxBindings } from "./create-docx.ts";
import { bindings as docxModelBindings } from './docx-model.ts';
import { bindings as appendRunBindings } from './append-run.ts';
import { bindings as rowHeaderBindings } from './row-header.ts';
import { bindings as paragraphTextBindings } from './paragraph-text.ts';
import { bindings as bodyInsertionBindings } from './body-insertion.ts';
import { bindings as bodyAnchorBindings, cleanupBodyAnchors } from './body-anchors.ts';
import { bindings as tableRowsBindings } from './table-rows.ts';
import { bindings as rowTextsBindings } from './row-texts.ts';
import { bindings as xmlByteSnapshotBindings } from './xml-byte-snapshot.ts';
import { bindings as formulaAnalysisBindings } from './formula-analysis.ts';
import { bindings as formulaRemapBindings } from './formula-remap.ts';
import { bindings as corePropertiesBindings } from './core-properties.ts';
import { bindings as documentPropertiesBindings } from './document-properties.ts';
import { bindings as tableStyleBindings } from './table-style.ts';
import { bindings as cellPropertiesBindings } from './cell-properties.ts';
import { bindings as createPptxBindings } from "./create-pptx.ts";
import { bindings as createXlsxBindings } from "./create-xlsx.ts";
import { bindings as tableDocxBindings } from "./tables-docx.ts";
import {bindings as nullableCellBindings} from './nullable-cell.ts';
import {bindings as tableMergingBindings,cleanupTableMerging} from './table-merging.ts';
import {bindings as verticalMergingBindings,cleanupVerticalMerging} from './vertical-merging.ts';
import {bindings as templateInventoryBindings,cleanupTemplateInventory} from './template-inventory.ts';
import {bindings as commentThreadBindings,cleanupCommentThreads} from './comment-threads.ts';
import {bindings as revisionPropertyBindings,cleanupRevisionProperties} from './revision-properties.ts';
import {bindings as revisionMoveBindings,cleanupRevisionMoves} from './revision-moves.ts';
import { bindings as tablePptxBindings } from "./tables-pptx.ts";
import { bindings as pptxBindings } from "./pptx.ts";
import { bindings as pptxCustodyBindings } from './pptx-custody.ts';
import { bindings as notesEditingBindings } from './notes-editing.ts';
import { bindings as xlsxBindings } from "./xlsx.ts";
import { bindings as xlsxRangeBindings } from './xlsx-range.ts';
import { bindings as xlsxCommentVmlBindings } from './xlsx-comment-vml.ts';
import { bindings as cacheBoundaryBindings } from "./cache-boundaries.ts";
import { bindings as workflowBindings, cleanupWorkflowFixtures } from "./workflow.ts";
import { bindings as workflowReceiptBindings, cleanupWorkflowReceipts } from './workflow-receipts.ts';
export async function cleanup(){await Promise.all([cleanupWorkflowFixtures(),cleanupWorkflowReceipts(),cleanupBodyAnchors(),cleanupTrackedWorkflowFixtures(),cleanupTrackingOutcomes(),cleanupTableMerging(),cleanupVerticalMerging(),cleanupTemplateInventory(),cleanupCommentThreads(),cleanupRevisionProperties(),cleanupRevisionMoves(),cleanupFontSize(),cleanupOpcCustody()]);}

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
  ...dataDescriptorBindings,
  ...namespaceBindings,
  ...docxBindings,
  ...storyBindings,
  ...revisionBindings,
  ...redlineBindings,
  ...commentBindings,
  ...formattingBindings,
  ...runEffectsBindings,
  ...runAppearanceBindings,
  ...runFontNameBindings,
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
  ...trackingSettingsBindings,
  ...trackingOutcomeBindings,
  ...createDocxBindings,
  ...docxModelBindings,
  ...appendRunBindings,
  ...rowHeaderBindings,
  ...paragraphTextBindings,
  ...bodyInsertionBindings,
  ...bodyAnchorBindings,
  ...tableRowsBindings,
  ...rowTextsBindings,
  ...xmlByteSnapshotBindings,
  ...formulaAnalysisBindings,
  ...formulaRemapBindings,
  ...corePropertiesBindings,
  ...documentPropertiesBindings,
  ...tableStyleBindings,
  ...cellPropertiesBindings,
  ...createPptxBindings,
  ...createXlsxBindings,
  ...tableDocxBindings,
  ...nullableCellBindings,
  ...tableMergingBindings,
  ...verticalMergingBindings,
  ...templateInventoryBindings,
  ...commentThreadBindings,
  ...revisionPropertyBindings,
  ...revisionMoveBindings,
  ...tablePptxBindings,
  ...pptxBindings,
  ...pptxCustodyBindings,
  ...notesEditingBindings,
  ...xlsxBindings,
  ...xlsxRangeBindings,
  ...xlsxCommentVmlBindings,
  ...cacheBoundaryBindings,
  ...workflowBindings,
  ...workflowReceiptBindings,
];

export default bindings;
