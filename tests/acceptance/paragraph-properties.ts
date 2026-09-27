import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { Document, type ParagraphAlignment } from '../../src/index.ts';
type State = { document: Document; error?: unknown };
const state = (c: Record<string, unknown>) => c.state as State;
export const scenarioIds = ['@id-docx-go-paragraph-alignment-getter', '@id-docx-go-paragraph-spacing-getters', '@id-docx-go-paragraph-advanced-toggles'];
export const bindings: StepBinding[] = [
  { pattern: /^a new Word paragraph$/, run: c => { const d = Document.create(); d.addParagraph(''); state(c).document = d; } },
  { pattern: /^its alignment is set to (\S+)$/, run: (c, value) => { state(c).document.paragraphs[0]!.setProperties({ alignment: value as ParagraphAlignment }); } },
  { pattern: /^its alignment getter equals (\S+)$/, run: (c, value) => { assert.equal(state(c).document.paragraphs[0]!.directProperties().alignment, value); } },
  { pattern: /^spacing before is set to (\d+) and after to (\d+)$/, run: (c, before, after) => { state(c).document.paragraphs[0]!.setProperties({ spacingBefore: Number(before), spacingAfter: Number(after) }); } },
  { pattern: /^its before and after getters equal (\d+) and (\d+)$/, run: (c, before, after) => { const values = state(c).document.paragraphs[0]!.directProperties(); assert.equal(values.spacingBefore, Number(before)); assert.equal(values.spacingAfter, Number(after)); } },
  { pattern: /^KeepLines, PageBreakBefore and WidowControl are set true$/, run: c => { state(c).document.paragraphs[0]!.setProperties({ keepLines: true, pageBreakBefore: true, widowControl: true }); } },
  { pattern: /^all three getters are true in memory$/, run: c => { const v = state(c).document.paragraphs[0]!.directProperties(); assert.equal(v.keepLines, true); assert.equal(v.pageBreakBefore, true); assert.equal(v.widowControl, true); } },
];
