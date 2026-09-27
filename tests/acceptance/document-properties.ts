import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
type State={document:Document};const state=(c:Record<string,unknown>)=>c.state as State;
export const scenarioIds=['@id-docx-go-section-title-background-getters'];
export const bindings:StepBinding[]=[
 {pattern:/^a new Word document with a first section$/,run:c=>{state(c).document=Document.create();}},
 {pattern:/^TitlePage is set true on that section and BackgroundColor to EEEEEE$/,run:c=>{assert.equal(state(c).document.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'}).changed,2);}},
 {pattern:/^the section TitlePage getter is (true|false) and the document BackgroundColor getter equals (\S+)$/,run:(c,flag,color)=>{const properties=state(c).document.getDocumentProperties();assert.equal(properties.titlePage,flag==='true');assert.equal(properties.backgroundColor,color);}},
];
