import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {xmlEquivalent} from '../../src/xml/comparison.ts';
const encoder=new TextEncoder();
type State={left?:Uint8Array;right?:Uint8Array;result?:boolean};
const state=(c:Record<string,unknown>)=>c.state as State;
export const bindings:StepBinding[]=[
 {pattern:/^the left XML is (.+)$/,run:(c,text)=>{state(c).left=encoder.encode(text!);}},
 {pattern:/^the right XML is (.+)$/,run:(c,text)=>{state(c).right=encoder.encode(text!);}},
 {pattern:/^the conservative XML comparator compares their UTF-8 bytes$/,run:c=>{const s=state(c);assert(s.left&&s.right);const l=s.left.slice(),r=s.right.slice();s.result=xmlEquivalent(s.left,s.right);assert.deepEqual(s.left,l);assert.deepEqual(s.right,r);}},
 {pattern:/^the comparison result is (true|false)$/,run:(c,expected)=>{assert.equal(state(c).result,expected==='true');}},
];
