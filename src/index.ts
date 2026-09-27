/** Bun-native OOXML entry points. See docs/agents/usage.md for safe edit workflows. */
export { OoxmlError } from "./errors.ts";
export { xmlEquivalent } from './xml/comparison.ts';
export { XmlSnapshot, type XmlRemovalTarget, type XmlAttributePatch, type XmlExpandedName, type XmlStructuredAttribute, type XmlContent, type XmlStructurePatch } from './xml/removal.ts';
export { admitPackage } from "./opc/admission.ts";
export { comparePackageArchives, type PackageComparison } from './opc/comparison.ts';
export type { ZipLimits } from "./opc/zip.ts";
export { OpcPackage, type PackageDiff, type Relationship } from "./opc/package.ts";
export { Document, Span, type AddParagraphOptions, type DirectCellProperties, type CellPropertiesPatch, type CellTopBorder, type DirectParagraphProperties, type ParagraphPropertiesPatch, type ParagraphAlignment, type RunFormattingPatch, type DirectRunFlags, type DirectRunAppearance, type UnderlineStyle, type HighlightColor, type RunVerticalAlignment, type RunFormattingReceipt, type EffectiveRunFormatting, type EffectiveFlag, type FormattingContribution, type AddParagraphStyleOptions, type ParagraphStyleDefinitionReceipt, type PageLayout } from "./docx/index.ts";
export { inspectStories, storyParts, type RevisionView, type StoryInspection } from "./docx/story.ts";
export { inspectRevisions, resolveRevisions, type Revision, type RevisionFinding } from "./docx/revisions.ts";
export { trackedReplace } from "./docx/redline.ts";
export { inspectComments, setCommentResolved, type CommentInfo, type CommentFinding, type CommentInspection } from "./docx/comments.ts";
export { Presentation, type NotesAnchor, type TextBoxGeometry, type TextBoxOptions, type TextBoxReceipt } from "./pptx/index.ts";
export { Workbook, parseA1Range, type A1Range, type A1Coordinate, type A1RangeAxis } from "./xlsx/index.ts";
export { patchOffice, type PatchRequest, type PatchReceipt, type TargetResult } from "./workflow/index.ts";
export type { Story, StoryKind } from "./docx/story.ts";
