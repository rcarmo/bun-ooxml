import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
type State={document:Document};
const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-docx-go-paragraph-text-getter'];
export const bindings:StepBinding[]=[
 {pattern:/^its text is set to JSON (.+)$/,run:(c,json)=>{const value:unknown=JSON.parse(json!);assert.equal(typeof value,'string');state(c).document.paragraphs[0]!.setText(value as string);}},
 {pattern:/^the paragraph text getter equals JSON (.+)$/,run:(c,json)=>{assert.equal(state(c).document.paragraphs[0]!.text,JSON.parse(json!));}},
];
