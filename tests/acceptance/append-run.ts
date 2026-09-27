import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {Document} from '../../src/index.ts';
import {parseXml,elements} from '../../src/xml/index.ts';
// The shared model save step performs actual path save/reopen and retains state.reopened.
type State={document:Document;reopened?:Document;before?:ReadonlyMap<string,Uint8Array>};
const state=(c:Record<string,unknown>)=>c.state as State;
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
export const scenarioIds=['@id-docx-go-paragraph-multiple-runs','@id-docx-go-roundtrip-selected-formatting'];
export const bindings:StepBinding[]=[
 {pattern:/^runs containing Hello-space, World and exclamation are appended in order$/,run:c=>{const s=state(c);for(const text of ['Hello ','World','!'])s.document.paragraphs[0]!.appendRun(text);}},
 {pattern:/^the paragraph has three runs and its text equals Hello World!$/,run:c=>{const p=state(c).document.paragraphs[0]!;assert.equal(p.directRunFlags().length,3);assert.equal(p.text,'Hello World!');}},
 {pattern:/^a new Word paragraph with three runs Bold-space, Italic-space and Colored$/,run:c=>{
  const s=state(c);s.document=Document.create();let p=s.document.addParagraph('');
  p=p.appendRun('Bold ',{bold:true});p=p.appendRun('Italic ',{italic:true});p=p.appendRun('Colored',{color:'FF0000',fontSizePt:14,fontName:'Arial'});
  assert.equal(p.text,'Bold Italic Colored');assert.equal(p.directRunFlags().length,3);
 }},
 {pattern:/^the first run is bold, the second italic, and the third has colour FF0000, font size 14 and font Arial$/,run:c=>{
  const s=state(c),p=s.document.paragraphs[0]!,flags=p.directRunFlags();
  assert.equal(flags[0]!.bold,true);assert.equal(flags[1]!.italic,true);assert.equal(p.directRunAppearance()[2]!.color,'FF0000');assert.equal(p.directFontSizes()[2],14);assert.equal(p.directFontNames()[2],'Arial');
  s.before=s.document.package.parts;
 }},
 {pattern:/^at least one paragraph and three runs are readable$/,run:c=>{
  const s=state(c);assert(s.reopened);assert(s.reopened.paragraphs.length>=1);assert(s.reopened.paragraphs[0]!.directRunFlags().length>=3);
 }},
 {pattern:/^the first run is bold and the second italic$/,run:c=>{const s=state(c);assert(s.reopened);const flags=s.reopened.paragraphs[0]!.directRunFlags();assert.equal(flags[0]!.bold,true);assert.equal(flags[1]!.italic,true);}},
 {pattern:/^the third run reports colour (\S+), font size (\d+) and font (.+)$/,run:(c,color,size,font)=>{
  const s=state(c);assert(s.reopened&&s.before);const p=s.reopened.paragraphs[0]!;
  assert.equal(p.directRunAppearance()[2]!.color,color);assert.equal(p.directFontSizes()[2],Number(size));assert.equal(p.directFontNames()[2],font);
  assert.equal(p.text,'Bold Italic Colored');const parsed=parseXml(new TextDecoder().decode(s.reopened.package.get('word/document.xml')));assert.equal(elements(parsed,'r',W).length,3);
  const after=s.reopened.package.parts;assert.deepEqual([...after.keys()],[...s.before.keys()]);for(const[n,b]of s.before)if(n!=='word/document.xml')assert.deepEqual(after.get(n),b);
 }},
];
