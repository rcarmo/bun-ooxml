import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {analyzeFormulaReferences,type FormulaReference,OoxmlError} from '../../src/index.ts';
type State={source?:string;refs?:FormulaReference[];error?:unknown};const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-xlsx-go-formula-analysis-counts','@id-xlsx-go-formula-quoted-sheet-flags','@id-xlsx-go-formula-analysis-refusal','@id-xlsx-go-formula-literal-punctuation'];
export const bindings:StepBinding[]=[
 {pattern:/^the formula source is JSON ("(?:[^"\\]|\\.)*")$/,run:(c,json)=>{const source:unknown=JSON.parse(json!);assert.equal(typeof source,'string');state(c).source=source as string;}},
 {pattern:/^the static formula analyser reads the source$/,run:c=>{const s=state(c);assert(s.source!==undefined);s.refs=undefined;s.error=undefined;try{s.refs=analyzeFormulaReferences(s.source);}catch(error){s.error=error;}}},
 {pattern:/^it returns (\d+) reference records without error$/,run:(c,count)=>{const s=state(c);assert.equal(s.error,undefined);assert(s.refs);assert.equal(s.refs.length,Number(count));}},
 {pattern:/^every reference has a nonempty byte span inside the original source$/,run:c=>{const s=state(c);assert(s.refs&&s.source!==undefined);const length=new TextEncoder().encode(s.source).length;for(const ref of s.refs){assert(Number.isInteger(ref.start)&&ref.start>=0);assert(Number.isInteger(ref.end)&&ref.end>ref.start&&ref.end<=length);}}},
 {pattern:/^its one reference spans every byte of the original source$/,run:c=>{const s=state(c);assert.equal(s.error,undefined);assert(s.refs&&s.source!==undefined);assert.equal(s.refs.length,1);assert.equal(s.refs[0]!.start,0);assert.equal(s.refs[0]!.end,new TextEncoder().encode(s.source).length);}},
 {pattern:/^the sheet is O'Brien with first cell column (\d+) row (\d+) and absolute column only$/,run:(c,col,row)=>{const r=state(c).refs?.[0];assert(r);assert.equal(r.sheet,"O'Brien");assert.equal(r.axis,'cell');assert.deepEqual(r.first,{column:Number(col),row:Number(row),columnAbsolute:true,rowAbsolute:false});}},
 {pattern:/^the last cell is column (\d+) row (\d+) with absolute row only$/,run:(c,col,row)=>{const r=state(c).refs?.[0];assert(r);assert.deepEqual(r.last,{column:Number(col),row:Number(row),columnAbsolute:false,rowAbsolute:true});}},
 {pattern:/^it returns an error and zero reference records$/,run:c=>{const s=state(c);assert(s.error instanceof OoxmlError);assert.equal(s.error.code,'xlsx-formula-unsupported');assert.equal(s.refs,undefined);}},
 {pattern:/^the analysis is (accepted without error|refused with zero references)$/,run:(c,result)=>{const s=state(c);if(result==='accepted without error'){assert.equal(s.error,undefined);assert(Array.isArray(s.refs));}else{assert(s.error instanceof OoxmlError);assert.equal(s.error.code,'xlsx-formula-unsupported');assert.equal(s.refs,undefined);}}},
];
