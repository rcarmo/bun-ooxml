import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document,type CoreProperties} from '../../src/index.ts';
type State={document:Document;core?:CoreProperties};const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-docx-go-core-properties-getters'];
export const bindings:StepBinding[]=[
 {pattern:/^its core properties are set to title Doc Title, creator Doc Author and subject Doc Subject$/,run:c=>{state(c).document.setCoreProperties({title:'Doc Title',creator:'Doc Author',subject:'Doc Subject'});}},
 {pattern:/^description Doc Description, keywords one;two, category Category and language en-US are supplied$/,run:c=>{state(c).document.setCoreProperties({description:'Doc Description',keywords:'one;two',category:'Category',language:'en-US'});}},
 {pattern:/^content status Draft, identifier urn:example:doc, last modifier Reviewer, revision 2 and version 1.0 are supplied$/,run:c=>{state(c).document.setCoreProperties({contentStatus:'Draft',identifier:'urn:example:doc',lastModifiedBy:'Reviewer',revision:'2',version:'1.0'});}},
 {pattern:/^created, modified and last-printed W3CDTF timestamps are supplied for 2026-02-03T00:00:00Z, 2026-02-03T01:00:00Z and 2026-02-03T02:00:00Z$/,run:c=>{state(c).document.setCoreProperties({created:'2026-02-03T00:00:00Z',modified:'2026-02-03T01:00:00Z',lastPrinted:'2026-02-03T02:00:00Z'});}},
 {pattern:/^the setter and getter return no error$/,run:c=>{state(c).core=state(c).document.getCoreProperties();assert(state(c).core);}},
 {pattern:/^only the in-memory title creator and subject are compared to (.+), (.+) and (.+)$/,run:(c,title,creator,subject)=>{const core=state(c).core;assert(core);assert.equal(core.title,title);assert.equal(core.creator,creator);assert.equal(core.subject,subject);}},
];
