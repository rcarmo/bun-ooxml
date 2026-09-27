import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
type State={document:Document;rowCounts?:number[];deleteError?:unknown};const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-docx-go-table-row-counts'];
export const bindings:StepBinding[]=[
 {pattern:/^a new Word table with two rows and three columns$/,run:c=>{const s=state(c);s.document=Document.create();s.document.addTable(2,3);}},
 {pattern:/^one row is appended, one is inserted at index one, and index one is deleted$/,run:c=>{const s=state(c);let t=s.document.tables[0]!;s.rowCounts=[];t=t.appendRow();s.rowCounts.push(t.rows);t=t.insertRow(1);s.rowCounts.push(t.rows);t=t.deleteRow(1);s.rowCounts.push(t.rows);}},
 {pattern:/^row counts after each step are (\w+), (\w+) and (\w+) respectively$/,run:(c,a,b,d)=>{const values:Record<string,number>={one:1,two:2,three:3,four:4};assert.deepEqual(state(c).rowCounts,[values[a!],values[b!],values[d!]]);}},
 {pattern:/^deletion at index ten returns an error$/,run:c=>{const s=state(c),before=s.document.package.toBytes();s.deleteError=undefined;try{s.document.tables[0]!.deleteRow(10);}catch(error){s.deleteError=error;}assert(s.deleteError instanceof RangeError);assert.deepEqual(s.document.package.toBytes(),before);}},
];
