import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {parseA1Range,type A1Range} from '../../src/xlsx/range.ts';
import {OoxmlError} from '../../src/errors.ts';
type State={source?:string;result?:A1Range;error?:unknown};
const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-xlsx-go-direct-range-parsing','@id-xlsx-go-direct-range-refusal'];
export const bindings:StepBinding[]=[
 {pattern:/^the direct range source is JSON (.+)$/,run:(c,json)=>{const source:unknown=JSON.parse(json!);assert.equal(typeof source,'string');state(c).source=source as string;}},
 {pattern:/^the direct-range parser reads the source$/,run:c=>{const s=state(c);assert(s.source!==undefined);try{s.result=parseA1Range(s.source);}catch(error){s.error=error;}}},
 {pattern:/^its sheet is JSON (.+) and its axis is (cell|row|column)$/,run:(c,sheet,axis)=>{const s=state(c);assert.equal(s.error,undefined);assert(s.result);assert.equal(s.result.sheet,JSON.parse(sheet!));assert.equal(s.result.axis,axis);}},
 {pattern:/^its first coordinate has row (\d+) and column (\d+)$/,run:(c,row,column)=>{const r=state(c).result;assert(r);assert.equal(r.first.row,Number(row));assert.equal(r.first.column,Number(column));}},
 {pattern:/^it returns an error$/,run:c=>{const s=state(c);assert(s.error instanceof OoxmlError);assert.equal(s.error.code,'xlsx-range-unsupported');assert.equal(s.result,undefined);}},
];
