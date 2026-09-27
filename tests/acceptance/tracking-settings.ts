import assert from 'node:assert/strict';
import {Document} from '../../src/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
export const scenarioIds=['@id-docx-go-track-author-toggle'];
const doc=(c:Record<string,unknown>)=>(c.state as {document:Document}).document;
export const bindings:StepBinding[]=[
 {pattern:/^a new Word document with tracking disabled$/,run:c=>{(c.state as {document:Document}).document=Document.create();assert.equal(doc(c).trackChangesEnabled,false);}},
 {pattern:/^tracking is enabled with Test Author and its author is changed to New Author$/,run:c=>{const d=doc(c);d.enableTracking('Test Author');assert.equal(d.trackAuthor,'Test Author');assert.equal(d.trackChangesEnabled,true);d.setTrackAuthor('New Author');}},
 {pattern:/^tracking is (enabled|disabled) and TrackAuthor equals (.+)$/,run:(c,enabled,author)=>{assert.equal(doc(c).trackChangesEnabled,enabled==='enabled');assert.equal(doc(c).trackAuthor,author);}},
 {pattern:/^disabling tracking makes TrackChangesEnabled (true|false)$/,run:(c,enabled)=>{doc(c).disableTracking();assert.equal(doc(c).trackChangesEnabled,enabled==='true');}},
];
