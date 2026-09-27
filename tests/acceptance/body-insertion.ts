import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
type State={document:Document;bodyCounts?:number[]};const state=(c:Record<string,unknown>)=>c.state as State;
const count=(d:Document)=>elements(parseXml(new TextDecoder().decode(d.package.get('word/document.xml'))),'body',W)[0]!.children.filter(n=>n.namespaceURI===W&&['p','tbl'].includes(n.localName)).length;
export const scenarioIds=['@id-docx-go-body-insert-order'];
export const bindings:StepBinding[]=[
 {pattern:/^a new Word body with no elements$/,run:c=>{const s=state(c);s.document=Document.create();assert.equal(count(s.document),0);}},
 {pattern:/^First and Third paragraphs are appended, then Second is inserted at index one$/,run:c=>{const s=state(c);s.bodyCounts=[];s.document.addParagraph('First');s.bodyCounts.push(count(s.document));s.document.addParagraph('Third');s.bodyCounts.push(count(s.document));s.document.insertParagraph(1,'Second');s.bodyCounts.push(count(s.document));}},
 {pattern:/^element counts after each operation are (\w+), (\w+) and (\w+)$/,run:(c,a,b,d)=>{const numbers:Record<string,number>={zero:0,one:1,two:2,three:3,four:4};assert.deepEqual(state(c).bodyCounts,[numbers[a!],numbers[b!],numbers[d!]]);}},
 {pattern:/^paragraph texts in order equal (.+), (.+) and (.+)$/,run:(c,a,b,d)=>{assert.deepEqual(state(c).document.paragraphs.map(p=>p.text),[a,b,d]);}},
];
