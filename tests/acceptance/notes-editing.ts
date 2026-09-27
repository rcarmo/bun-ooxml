import assert from 'node:assert/strict';
import {mkdtemp,rm} from 'node:fs/promises';
import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Presentation,type NotesAnchor} from '../../src/pptx/index.ts';
import {OpcPackage} from '../../src/opc/package.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
import {fixturePath,F} from '../../scripts/fixture-inputs.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
const N='ppt/notesSlides/notesSlide1.xml',A='http://schemas.openxmlformats.org/drawingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main';
const ID='fixture-04faba67841dda25dc3ff9e3e6e345e6feeeef1cf25a6b9065bf5fbdc83163dc';
type State={deck:Presentation;original:Uint8Array;xml:string;target:NotesAnchor;other?:NotesAnchor;error?:unknown;saved?:Uint8Array;receipt?:{changedParts:string[]};partBefore?:Uint8Array;foreignDeck?:Presentation;refusals?:unknown[]};
const state=(c:Record<string,unknown>)=>c.state as State;
async function prepare(s:State,paragraph?:string){s.original=await Bun.file(fixturePath(F.goSlides.notes)).bytes();const p=await OpcPackage.open(s.original);if(paragraph){const original='<a:p><a:r><a:t>Remember to emphasize the Gothic elements</a:t></a:r></a:p>';assert(p.text(N).includes(original));p.set(N,p.text(N).replace(original,paragraph));}s.deck=await Presentation.open(p.toBytes());s.xml=s.deck.package.text(N);s.target=s.deck.slides[0]!.inspectNotesText();}
const replace=(s:State,text:string)=>s.receipt=s.deck.slides[0]!.replaceNotesAt(s.target,text);
const read=(s:State)=>s.deck.slides[0]!.inspectNotesText().text;
export const bindings:StepBinding[]=[
 {pattern:new RegExp('^two presentation edit sessions open fixture '+ID+'$'),run:async c=>{const s=state(c);await prepare(s);s.foreignDeck=await Presentation.open(s.original);s.refusals=[];}},
 {pattern:/^the first session holds notes for ppt\/slides\/slide1.xml$/,run:c=>{const s=state(c);s.target=s.deck.slides[0]!.inspectNotesText();assert.equal(s.target.text,'Remember to emphasize the Gothic elements');}},
 {pattern:/^the other session attempts to write foreign through that target$/,run:c=>{const s=state(c);let error:unknown;try{s.foreignDeck!.slides[0]!.replaceNotesAt(s.target,'foreign');}catch(e){error=e;}s.refusals!.push(error);}},
 {pattern:/^the first session attempts JSON notes text (".*"), one invalid UTF-8 byte FF, and JSON text (".*")$/,run:(c,nul,tab)=>{const s=state(c),slide=s.deck.slides[0]!;for(const write of [()=>slide.replaceNotesAt(s.target,JSON.parse(nul!)),()=>slide.replaceNotesUtf8At(s.target,new Uint8Array([0xff])),()=>slide.replaceNotesAt(s.target,JSON.parse(tab!))]){let error:unknown;try{write();}catch(e){error=e;}s.refusals!.push(error);}}},
 {pattern:/^each foreign or invalid write returns an error$/,run:c=>{const s=state(c);assert.equal(s.refusals!.length,4);for(const error of s.refusals!)assert(error instanceof Error);assert.equal((s.refusals![2] as {code?:string}).code,'PPTX_NOTES_UTF8_INVALID');assert.deepEqual(s.deck.package.toBytes(),s.original);assert.deepEqual(s.foreignDeck!.package.toBytes(),s.original);}},
 {pattern:/^replacing through the held target with its unchanged text succeeds$/,run:c=>{const s=state(c);assert.deepEqual(replace(s,s.target.text),{changedParts:[]});assert.equal(read(s),s.target.text);}},
 {pattern:/^serialising the first session at that point gives the exact original fixture archive bytes$/,run:c=>{const s=state(c);assert.deepEqual(s.deck.package.toBytes(),s.original);}},
 {pattern:new RegExp('^fixture '+ID+'$'),run:async c=>prepare(state(c))},
 {pattern:/^the notes for ppt\/slides\/slide1.xml read Remember to emphasize the Gothic elements$/,run:c=>assert.equal(read(state(c)),'Remember to emphasize the Gothic elements')},
 {pattern:/^those notes are replaced with Updated speaker notes and saved to a new PPTX path$/,run:async c=>{const s=state(c);replace(s,'Updated speaker notes');const root=await mkdtemp(join(tmpdir(),'notes-edit-'));try{const path=join(root,'notes.pptx');await s.deck.save(path);s.saved=await Bun.file(path).bytes();}finally{await rm(root,{recursive:true,force:true});}}},
 {pattern:/^the in-memory ppt\/notesSlides\/notesSlide1.xml payload equals one replacement of the original speaker text$/,run:c=>{const s=state(c);assert.equal(s.deck.package.text(N),s.xml.replace('Remember to emphasize the Gothic elements','Updated speaker notes'));}},
 {pattern:/^the save receipt lists exactly ppt\/notesSlides\/notesSlide1.xml as its only changed part$/,run:c=>assert.deepEqual(state(c).receipt,{changedParts:[N]})},
 {pattern:/^every other part in the original package graph has the same delivered payload bytes$/,run:async c=>{const s=state(c),before=await OpcPackage.open(s.original),after=await OpcPackage.open(s.saved!);assert.deepEqual(before.names(),after.names());for(const n of before.names())if(n!==N)assert.deepEqual(before.get(n),after.get(n));}},
 {pattern:/^reopening the saved PPTX reads Updated speaker notes for ppt\/slides\/slide1.xml$/,run:async c=>{assert.equal((await Presentation.open(state(c).saved!)).slides[0]!.inspectNotesText().text,'Updated speaker notes');}},
 {pattern:new RegExp('^a presentation edit session holds a notes target for ppt/slides/slide1.xml from fixture '+ID+'$'),run:async c=>prepare(state(c))},
 {pattern:/^a second notes target is found in the same session before clearing$/,run:c=>{const s=state(c);s.other=s.deck.slides[0]!.inspectNotesText();}},
 {pattern:/^the first target replaces its notes with the empty string$/,run:c=>{replace(state(c),'');}},
 {pattern:/^replacement through the second held target with stale returns an error$/,run:c=>{const s=state(c),before=s.deck.package.toBytes();assert.throws(()=>s.deck.slides[0]!.replaceNotesAt(s.other!,'stale'));assert.deepEqual(s.deck.package.toBytes(),before);}},
 {pattern:/^a newly found notes target reads the empty string$/,run:c=>{const s=state(c);s.target=s.deck.slides[0]!.inspectNotesText();assert.equal(s.target.text,'');}},
 {pattern:/^replacing through that fresh target with refilled succeeds$/,run:c=>{replace(state(c),'refilled');assert.equal(read(state(c)),'refilled');}},
 {pattern:new RegExp('^a presentation edit session holds notes for ppt/slides/slide1.xml from fixture '+ID+'$'),run:async c=>prepare(state(c))},
 {pattern:/^the session's retained ppt\/notesSlides\/notesSlide1.xml payload is directly replaced by changing Gothic to Victorian$/,run:c=>{const s=state(c);s.deck.package.set(N,s.xml.replace('Gothic','Victorian'));}},
 {pattern:/^the older target attempts to replace notes with stale$/,run:c=>{try{replace(state(c),'stale');}catch(error){state(c).error=error;}}},
 {pattern:/^that notes replacement returns an error$/,run:c=>assert(state(c).error instanceof Error)},
 {pattern:/^the pinned notes fixture's first notes paragraph is replaced in memory by (.+)$/,run:async(c,paragraph)=>prepare(state(c),paragraph)},
 {pattern:/^the editor replaces its notes with filled$/,run:c=>{replace(state(c),'filled');}},
 {pattern:/^a newly found notes target for ppt\/slides\/slide1.xml reads filled$/,run:c=>assert.equal(read(state(c)),'filled')},
 {pattern:/^the editor replaces its notes with JSON (.+)$/,run:(c,json)=>{replace(state(c),JSON.parse(json!));}},
 {pattern:/^the parsed ppt\/notesSlides\/notesSlide1.xml has a DrawingML text leaf equal to JSON (.+)$/,run:(c,json)=>assert(elements(parseXml(state(c).deck.package.text(N)),'t',A).some(t=>t.text===JSON.parse(json!)))},
 {pattern:/^that same text leaf has an XML-namespace space attribute equal to preserve$/,run:c=>{const t=elements(parseXml(state(c).deck.package.text(N)),'t',A).find(t=>t.text===' leading ');assert(t);assert.equal(attribute(t,'space','http://www.w3.org/XML/1998/namespace'),'preserve');}},
 {pattern:/^an identical-text replacement is applied to that notes target$/,run:c=>{const s=state(c);s.partBefore=s.deck.package.get(N);replace(s,s.target.text);}},
 {pattern:/^the notes part payload bytes are unchanged$/,run:c=>assert.deepEqual(state(c).deck.package.get(N),state(c).partBefore)},
 {pattern:/^the editor replaces through that target with JSON (.+)$/,run:(c,json)=>{replace(state(c),JSON.parse(json!));}},
 {pattern:/^a newly found notes target reads JSON (.+) across four paragraphs$/,run:(c,json)=>{const s=state(c);assert.equal(read(s),JSON.parse(json!));const doc=parseXml(s.deck.package.text(N)),body=elements(doc,'sp',P).find(n=>elements(n,'ph',P).some(ph=>attribute(ph,'type')==='body'))!;assert.equal(elements(body,'p',A).length,4);}},
 {pattern:/^the complete notes part has no literal b="1" attribute$/,run:c=>assert(!state(c).deck.package.text(N).includes('b="1"'))},
 {pattern:/^the notes XML contains exactly two copies each of val="112233", typeface="F&amp;F", lang="en-US" and char="•"$/,run:c=>{for(const token of ['val="112233"','typeface="F&amp;F"','lang="en-US"','char="•"'])assert.equal(state(c).deck.package.text(N).split(token).length-1,2);}},
];
export const scenarioIds=['@id-pptx-go-notes-refusal-and-noop-custody','@id-pptx-go-notes-exact-splice','@id-pptx-go-notes-clear-stale-refill','@id-pptx-go-notes-part-fingerprint-refusal','@id-pptx-go-notes-self-closing-fill','@id-pptx-go-notes-edge-space-preserve','@id-pptx-go-notes-multiline-template','@id-pptx-go-notes-template-fragments'];
