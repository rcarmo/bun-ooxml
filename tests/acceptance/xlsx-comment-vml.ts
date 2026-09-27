import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {fixturePath} from '../../scripts/fixture-inputs.ts';
import {OpcPackage,type Relationship} from '../../src/opc/package.ts';
import {parseXml,attribute,elements} from '../../src/xml/index.ts';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ID='fixture-264be55e012d4bc2b3bf25e59824fdd30022e94f70869ea7d6ad960b803a902f',sheet='xl/worksheets/sheet1.xml';
export type CommentVmlState={input:Uint8Array;original:Uint8Array;pkg:OpcPackage;legacyId?:string;relationships?:Relationship[];vml?:Relationship;comments?:Relationship;commentRows?:{ref:string|undefined;text:string}[]};
const state=(c:Record<string,unknown>)=>c.state as CommentVmlState;
const custody=(s:CommentVmlState)=>{assert.deepEqual(s.input,s.original);assert.deepEqual(s.pkg.toBytes(),s.original);};
export const scenarioIds=['@id-xlsx-comment-vml-existing-graph'];
export const bindings:StepBinding[]=[
 {pattern:new RegExp('^fixture '+ID+'$'),run:async c=>{const s=state(c);s.input=await Bun.file(fixturePath(ID)).bytes();s.original=s.input.slice();s.pkg=await OpcPackage.open(s.input);}},
 {pattern:/^a namespace-aware reader opens xl\/worksheets\/sheet1.xml and its relationship part$/,run:c=>{
  const s=state(c),doc=parseXml(s.pkg.text(sheet));assert.equal(doc.root.localName,'worksheet');assert.equal(doc.root.namespaceURI,S);
  const drawings=doc.root.children.filter(n=>n.localName==='legacyDrawing'&&n.namespaceURI===S);assert.equal(drawings.length,1);s.legacyId=attribute(drawings[0]!,'id',R);assert(s.legacyId);
  s.relationships=s.pkg.relationships(sheet);const vml=s.relationships.filter(r=>r.id===s.legacyId);assert.equal(vml.length,1);s.vml=vml[0]!;
  const comments=s.relationships.filter(r=>r.type===R+'/comments');assert.equal(comments.length,1);s.comments=comments[0]!;assert.equal(s.comments.external,false);assert(s.comments.resolved);
  const content=parseXml(s.pkg.text(s.comments.resolved));assert.equal(content.root.localName,'comments');assert.equal(content.root.namespaceURI,S);const lists=content.root.children.filter(n=>n.localName==='commentList'&&n.namespaceURI===S);assert.equal(lists.length,1);
  s.commentRows=lists[0]!.children.map(n=>{assert.equal(n.localName,'comment');assert.equal(n.namespaceURI,S);const text=n.children.filter(n=>n.localName==='text'&&n.namespaceURI===S);assert.equal(text.length,1);return {ref:attribute(n,'ref'),text:elements(text[0]!,'t',S).map(n=>n.text).join('')};});
  custody(s);
 }},
 {pattern:/^the legacyDrawing relationship ID is (\S+) and resolves internally to (\S+)$/,run:(c,id,target)=>{const s=state(c);assert.equal(s.legacyId,id);assert.equal(s.vml!.id,id);assert.equal(s.vml!.external,false);assert.equal(s.vml!.resolved,target);assert(s.pkg.get(target!));custody(s);}},
 {pattern:/^the corresponding relationship Type is (\S+)$/,run:(c,type)=>{const s=state(c);assert.equal(s.vml!.type,type);custody(s);}},
 {pattern:/^a separate comments relationship resolves internally to (\S+)$/,run:(c,target)=>{const s=state(c);assert.equal(s.comments!.type,R+'/comments');assert.equal(s.comments!.external,false);assert.equal(s.comments!.resolved,target);assert.notEqual(s.comments!.id,s.vml!.id);assert.notEqual(s.comments!.resolved,s.vml!.resolved);assert(s.pkg.get(target!));custody(s);}},
 {pattern:/^the comment part contains (\S+) with text (.+)$/,run:(c,ref,text)=>{const s=state(c),rows=s.commentRows!.filter(row=>row.ref===ref);assert.equal(rows.length,1);assert.equal(rows[0]!.text,text);custody(s);}},
];
