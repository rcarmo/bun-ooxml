import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
type State={document:Document};const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-docx-go-table-cell-text-getters'];
export const bindings:StepBinding[]=[
 {pattern:/^a new Word table with two rows and two columns$/,run:c=>{const s=state(c);s.document=Document.create();s.document.addTable(2,2);}},
 {pattern:/^its cells are set by row to A1, B1, A2 and B2$/,run:c=>{const t=state(c).document.tables[0]!;for(const[i,value]of ['A1','B1','A2','B2'].entries())t.cell(Math.floor(i/2),i%2).text=value;}},
 {pattern:/^the four cell text getters equal (\S+), (\S+), (\S+) and (\S+) in those positions$/,run:(c,...values)=>{const t=state(c).document.tables[0]!;assert.equal(values.length,4);for(const[i,value]of values.entries())assert.equal(t.cell(Math.floor(i/2),i%2).text,value);}},
 {pattern:/^FirstRowText returns exactly (\S+) and (\S+)$/,run:(c,a,b)=>{assert.deepEqual(state(c).document.tables[0]!.rowTexts(0),[a,b]);}},
];
