/** Bun-native OOXML entry points. See docs/agents/usage.md for safe edit workflows. */
export { OoxmlError } from "./errors.ts";
export { OpcPackage, type PackageDiff, type Relationship } from "./opc/package.ts";
export { Document, Span, type RunFormattingPatch, type RunFormattingReceipt, type AddParagraphStyleOptions, type ParagraphStyleDefinitionReceipt } from "./docx/index.ts";
export { inspectStories, storyParts, type RevisionView, type StoryInspection } from "./docx/story.ts";
export { inspectRevisions, resolveRevisions, type Revision, type RevisionFinding } from "./docx/revisions.ts";
export { trackedReplace } from "./docx/redline.ts";
export { inspectComments, setCommentResolved, type CommentInfo, type CommentFinding, type CommentInspection } from "./docx/comments.ts";
export { Presentation, type TextBoxGeometry, type TextBoxOptions, type TextBoxReceipt } from "./pptx/index.ts";
export { Workbook } from "./xlsx/index.ts";
export { patchOffice, type PatchRequest, type PatchReceipt, type TargetResult } from "./workflow/index.ts";
export type { Story, StoryKind } from "./docx/story.ts";
