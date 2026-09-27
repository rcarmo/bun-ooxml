import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {XmlByteSnapshot} from '../../src/index.ts';
type State={source:string;input?:Uint8Array;before?:Uint8Array;byteSnapshot?:XmlByteSnapshot;output?:Uint8Array};const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-xml-go-immutable-leaf-seed'];
export const bindings:StepBinding[]=[
 // The existing XML-source Given owns s.source and a separate string snapshot.
 {pattern:/^Go parses a caller-owned byte slice and performs an empty edit$/,run:c=>{const s=state(c);assert.equal(typeof s.source,'string');s.input=new TextEncoder().encode(s.source);s.before=new Uint8Array(s.input);s.byteSnapshot=XmlByteSnapshot.parse(s.input);s.output=s.byteSnapshot.remove([]);}},
 {pattern:/^the caller input bytes still equal the original XML source$/,run:c=>{const s=state(c);assert(s.input&&s.before);assert.deepEqual(s.input,s.before);assert.deepEqual(s.input,new TextEncoder().encode(s.source));}},
 {pattern:/^the empty edit returns the exact original source bytes$/,run:c=>{const s=state(c);assert(s.output&&s.before);assert.deepEqual(s.output,s.before);assert.deepEqual(s.output,new TextEncoder().encode(s.source));}},
];
