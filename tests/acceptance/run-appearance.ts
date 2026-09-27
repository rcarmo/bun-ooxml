import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { Document, type UnderlineStyle, type HighlightColor } from '../../src/index.ts';
type State={document:Document};const state=(c:Record<string,unknown>)=>c.state as State;
function runDocument(){const document=Document.create();document.addParagraph('Test');return document;}
const paragraph=(s:State)=>s.document.paragraphs[0]!;
export const scenarioIds=['@id-docx-go-run-color-getter','@id-docx-go-run-underline-style','@id-docx-go-run-highlight','@id-docx-go-run-vertical-align'];
export const bindings:StepBinding[]=[
 {pattern:/^a new Word run$/,run:c=>{state(c).document=runDocument();}},
 {pattern:/^its colour is set to (\S+)$/,run:(c,value)=>{paragraph(state(c)).setRunFormatting({color:value!});}},
 {pattern:/^its in-memory colour getter equals (\S+)$/,run:(c,value)=>{const rows=paragraph(state(c)).directRunAppearance();assert.equal(rows.length,1);assert.equal(rows[0]!.color,value);}},
 {pattern:/^its underline style is set to (\S+)$/,run:(c,value)=>{paragraph(state(c)).setRunFormatting({underline:value as UnderlineStyle});}},
 {pattern:/^Underline is true and UnderlineStyle equals (\S+)$/,run:(c,value)=>{const rows=paragraph(state(c)).directRunAppearance();assert.equal(rows.length,1);assert(rows[0]!.underline!==null&&rows[0]!.underline!=='none');assert.equal(rows[0]!.underline,value);}},
 {pattern:/^highlight is set to (\S+)$/,run:(c,value)=>{paragraph(state(c)).setRunFormatting({highlight:value as HighlightColor});}},
 {pattern:/^the Highlight getter equals (\S+)$/,run:(c,value)=>{const rows=paragraph(state(c)).directRunAppearance();assert.equal(rows.length,1);assert.equal(rows[0]!.highlight,value);}},
 {pattern:/^two new Word runs$/,run:c=>{const s=state(c);s.document=runDocument();s.document.addParagraph('Test');}},
 {pattern:/^Superscript is enabled on the first and Subscript on the second$/,run:c=>{const s=state(c);paragraph(s).setRunFormatting({verticalAlign:'superscript'});s.document.paragraphs[1]!.setRunFormatting({verticalAlign:'subscript'});}},
 {pattern:/^the first reports superscript true and subscript false$/,run:c=>{const rows=paragraph(state(c)).directRunAppearance();assert.equal(rows.length,1);assert.equal(rows[0]!.verticalAlign,'superscript');assert.notEqual(rows[0]!.verticalAlign,'subscript');}},
 {pattern:/^the second reports subscript true and superscript false$/,run:c=>{const s=state(c);const rows=s.document.paragraphs[1]!.directRunAppearance();assert.equal(rows.length,1);assert.equal(rows[0]!.verticalAlign,'subscript');assert.notEqual(rows[0]!.verticalAlign,'superscript');}},
];
