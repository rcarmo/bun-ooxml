import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {fixturePath} from '../../scripts/fixture-inputs.ts';
import {OpcPackage,type Relationship} from '../../src/opc/package.ts';
import {inspectWorksheetComments} from '../../src/xlsx/comments.ts';
const R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const ID='fixture-264be55e012d4bc2b3bf25e59824fdd30022e94f70869ea7d6ad960b803a902f',sheet='xl/worksheets/sheet1.xml';
export type CommentVmlState={input:Uint8Array;original:Uint8Array;pkg:OpcPackage;legacyId?:string;relationships?:Relationship[];vml?:Relationship;comments?:Relationship;commentRows?:{ref:string|undefined;text:string}[]};
const state=(c:Record<string,unknown>)=>c.state as CommentVmlState;
const custody=(s:CommentVmlState)=>{assert.deepEqual(s.input,s.original);assert.deepEqual(s.pkg.toBytes(),s.original);};
export const scenarioIds=['@id-xlsx-comment-vml-existing-graph'];
export const bindings:StepBinding[]=[
 {pattern:new RegExp('^fixture '+ID+'$'),run:async c=>{const s=state(c);s.input=await Bun.file(fixturePath(ID)).bytes();s.original=s.input.slice();s.pkg=await OpcPackage.open(s.input);}},
 {pattern:/^a namespace-aware reader opens xl\/worksheets\/sheet1.xml and its relationship part$/,run:c=>{
  const s=state(c),result=inspectWorksheetComments(s.pkg,sheet);assert(result.vmlRelationship);assert(result.commentsRelationship);
  s.legacyId=result.vmlRelationship.id;
  s.vml={id:result.vmlRelationship.id,type:result.vmlRelationship.type,target:result.vmlRelationship.part,resolved:result.vmlRelationship.part,external:false};
  s.comments={id:result.commentsRelationship.id,type:result.commentsRelationship.type,target:result.commentsRelationship.part,resolved:result.commentsRelationship.part,external:false};
  s.commentRows=result.comments.map(row=>({ref:row.reference,text:row.text}));
  custody(s);
 }},
 {pattern:/^the legacyDrawing relationship ID is (\S+) and resolves internally to (\S+)$/,run:(c,id,target)=>{const s=state(c);assert.equal(s.legacyId,id);assert.equal(s.vml!.id,id);assert.equal(s.vml!.external,false);assert.equal(s.vml!.resolved,target);assert(s.pkg.get(target!));custody(s);}},
 {pattern:/^the corresponding relationship Type is (\S+)$/,run:(c,type)=>{const s=state(c);assert.equal(s.vml!.type,type);custody(s);}},
 {pattern:/^a separate comments relationship resolves internally to (\S+)$/,run:(c,target)=>{const s=state(c);assert.equal(s.comments!.type,R+'/comments');assert.equal(s.comments!.external,false);assert.equal(s.comments!.resolved,target);assert.notEqual(s.comments!.id,s.vml!.id);assert.notEqual(s.comments!.resolved,s.vml!.resolved);assert(s.pkg.get(target!));custody(s);}},
 {pattern:/^the comment part contains (\S+) with text (.+)$/,run:(c,ref,text)=>{const s=state(c),rows=s.commentRows!.filter(row=>row.ref===ref);assert.equal(rows.length,1);assert.equal(rows[0]!.text,text);custody(s);}},
];
