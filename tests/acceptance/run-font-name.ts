import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import type { Document } from '../../src/index.ts';
const document=(c:Record<string,unknown>)=>(c.state as {document:Document}).document;
export const scenarioIds=['@id-docx-go-run-font-name'];
export const bindings:StepBinding[]=[
 {pattern:/^its font name is set to (.+)$/,run:(c,name)=>{document(c).paragraphs[0]!.setRunFormatting({fontName:name!});}},
 {pattern:/^its font-name getter equals (.+)$/,run:(c,name)=>{const names=document(c).paragraphs[0]!.directFontNames();assert.equal(names.length,1);assert.equal(names[0],name);}},
];
