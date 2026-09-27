import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { Document } from '../../src/index.ts';
type State = { document: Document };
const state = (c: Record<string, unknown>) => c.state as State;
export const scenarioIds = ['@id-docx-go-run-boolean-formatting'];
export const bindings: StepBinding[] = [
  { pattern: /^a new Go Word run containing Test$/, run: c => { const document = Document.create(); document.addParagraph('Test'); state(c).document = document; } },
  { pattern: /^bold is set to (true|false), italic to (true|false) and strike to (true|false)$/, run: (c, bold, italic, strike) => {
    state(c).document.paragraphs[0]!.setRunFormatting({ bold: bold === 'true', italic: italic === 'true', strike: strike === 'true' });
  } },
  { pattern: /^Bold, Italic and Strike getters equal (true|false), (true|false) and (true|false)$/, run: (c, bold, italic, strike) => {
    const flags = state(c).document.paragraphs[0]!.directRunFlags(); assert.equal(flags.length, 1);
    assert.equal(flags[0]!.bold, bold === 'true'); assert.equal(flags[0]!.italic, italic === 'true'); assert.equal(flags[0]!.strike, strike === 'true');
  } },
];
