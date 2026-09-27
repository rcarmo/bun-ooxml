import assert from 'node:assert/strict';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
import {Document,type BodyAnchor} from '../../src/index.ts';
import type {StepBinding} from '../../scripts/gherkin.ts';
type State={root:string;path:string;document:Document;anchors:readonly BodyAnchor[];selected?:BodyAnchor;inserted?:string;success?:boolean;original:Uint8Array};
const state=(c:Record<string,unknown>)=>c.state as State,roots=new Set<string>();
export async function cleanupBodyAnchors(){await Promise.all([...roots].map(async root=>{await rm(root,{recursive:true,force:true});roots.delete(root);}));}
async function setup(c:Record<string,unknown>,first:string,second:string,body1:string,body2:string){
 const s=state(c);s.root=await mkdtemp(join(tmpdir(),'body-anchor-acceptance-'));roots.add(s.root);s.path=join(s.root,'document.docx');
 const d=Document.create();d.addParagraph(first).setProperties({outlineLevel:0});d.addParagraph(body1);d.addParagraph(second).setProperties({outlineLevel:0});d.addParagraph(body2);await d.save(s.path);s.original=await Bun.file(s.path).bytes();s.document=await Document.open(s.path);
}
export const scenarioIds=['@id-python-word-anchor-headings-paragraphs','@id-python-word-anchor-text-filter','@id-python-word-anchor-discover-insert'];
export const bindings:StepBinding[]=[
 {pattern:/^a saved Word document has headings "([^"]+)" and "([^"]+)" and paragraphs "([^"]+)" and "([^"]+)"$/,run:setup},
 {pattern:/^a saved Word document has headings "([^"]+)" and "([^"]+)" with paragraphs "([^"]+)" and "([^"]+)"$/,run:setup},
 {pattern:/^Word anchors are listed without a query$/,run:async c=>{const s=state(c);s.anchors=s.document.inspectBodyAnchors();assert.deepEqual(s.document.package.toBytes(),s.original);assert.deepEqual(await Bun.file(s.path).bytes(),s.original);}},
 {pattern:/^Word anchors are listed with query "([^"]+)"$/,run:async(c,query)=>{const s=state(c);s.anchors=s.document.inspectBodyAnchors(query);assert.deepEqual(s.document.package.toBytes(),s.original);assert.deepEqual(await Bun.file(s.path).bytes(),s.original);}},
 {pattern:/^the anchor count is at least (\d+)$/,run:(c,n)=>assert(state(c).anchors.length>=Number(n))},
 {pattern:/^an anchor has type "([^"]+)" and text "([^"]+)"$/,run:(c,type,text)=>assert(state(c).anchors.some(a=>a.type===type&&a.text===text))},
 {pattern:/^a "([^"]+)" anchor contains "([^"]+)" in its text$/,run:(c,type,text)=>assert(state(c).anchors.some(a=>a.type===type&&a.text.includes(text!)))},
 {pattern:/^every returned anchor text contains "([^"]+)" case-insensitively$/,run:(c,query)=>{const anchors=state(c).anchors;assert(anchors.length>0);assert(anchors.every(a=>a.text.toLowerCase().includes(query!.toLowerCase())));}},
 {pattern:/^the anchor whose text equals "([^"]+)" is selected$/,run:(c,text)=>{const s=state(c),found=s.anchors.filter(a=>a.text===text);assert.equal(found.length,1);s.selected=found[0]!;}},
 {pattern:/^"([^"]+)" is inserted after that anchor in the source document$/,run:async(c,text)=>{const s=state(c);assert(s.selected);const p=s.document.insertParagraphAfter(s.selected,text!);assert.equal(p.text,text);await s.document.save(s.path);s.inserted=text;s.success=true;}},
 {pattern:/^the insertion response has success true$/,run:c=>assert.equal(state(c).success,true)},
 {pattern:/^reading the saved Word document shows "([^"]+)" immediately after "([^"]+)"$/,run:async(c,text,heading)=>{const s=state(c),d=await Document.open(s.path),rows=d.paragraphs.map(p=>p.text),index=rows.indexOf(heading!);assert(index>=0);assert.equal(rows[index+1],text);}},
];
