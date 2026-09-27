import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
type State={document:Document;styleBefore?:string;styleAfter?:string};const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-docx-go-table-style-getter'];
export const bindings:StepBinding[]=[
 {pattern:/^a new Word two-by-two table$/,run:c=>{const s=state(c);s.document=Document.create();s.document.addTable(2,2);}},
 {pattern:/^its style is read, then set to TableGrid and read again$/,run:c=>{const s=state(c),t=s.document.tables[0]!;s.styleBefore=t.styleId??'';t.setStyle('TableGrid');s.styleAfter=t.styleId??'';}},
 {pattern:/^the first style is empty and the second style is (\S+)$/,run:(c,value)=>{const s=state(c);assert.equal(s.styleBefore,'');assert.equal(s.styleAfter,value);}},
];
