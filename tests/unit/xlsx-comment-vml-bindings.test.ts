import {test,expect} from 'bun:test';
import {join} from 'node:path';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {executeAcceptance,selectSharedScenarios,type StepBinding} from '../../scripts/gherkin.ts';
import {bindings} from '../acceptance/steps.ts';
import type {CommentVmlState} from '../acceptance/xlsx-comment-vml.ts';
const path='workflows/xlsx/comment-vml-custody.feature',id='@id-xlsx-comment-vml-existing-graph';
const given='fixture fixture-264be55e012d4bc2b3bf25e59824fdd30022e94f70869ea7d6ad960b803a902f';
const read='a namespace-aware reader opens xl/worksheets/sheet1.xml and its relationship part';
const sheet='xl/worksheets/sheet1.xml',rels='xl/worksheets/_rels/sheet1.xml.rels',comments='xl/comments/comment1.xml',vml='xl/drawings/commentsDrawing1.vml';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
async function run(change:(text:string)=>string=text=>text,custom:StepBinding[]=bindings){
 const feature=selectSharedScenarios(path,change(await Bun.file(join(fixturesRoot(),path)).text()),[id]),count=(n:number)=>({implemented:n,planned:0,total:n});
 return executeAcceptance({root:'.',features:[feature],counts:{features:count(1),scenarios:count(1),cases:count(1),steps:count(7)}},custom,'comment-vml-unit');
}
test('canonical XLSX comment and VML inspection checks both relationships and exact comment values',async()=>{
 const result=await run();expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(1);expect(result.counts.cases.planned).toBe(5);expect(result.counts.steps.passed).toBe(7);expect(result.counts.steps.undefined).toBe(0);expect(result.counts.steps.ambiguous).toBe(0);
});

function changed(step:string,mutate:(s:CommentVmlState)=>void){expect(bindings.filter(b=>b.pattern.test(step))).toHaveLength(1);return bindings.map(b=>b.pattern.test(step)?{...b,run:async(c:Record<string,unknown>,...args:string[])=>{await b.run(c,...args);mutate(c.state as CommentVmlState);}}:b);}
function fails(r:Awaited<ReturnType<typeof run>>){expect(r.counts.cases.failed).toBe(1);expect(r.counts.steps.failed).toBe(1);expect(r.counts.steps.undefined).toBe(0);expect(r.counts.steps.ambiguous).toBe(0);}
function baseline(s:CommentVmlState){s.original=s.pkg.toBytes();s.input=s.original.slice();}

test('each canonical relationship identity type target and comment text rejects a wrong expected value',async()=>{
 for(const [from,to]of [
  ['ID is anysvml','ID is wrong'],
  ['internally to xl/drawings/commentsDrawing1.vml','internally to xl/drawings/wrong.vml'],
  ['Type is '+R+'/vmlDrawing','Type is '+R+'/comments'],
  ['internally to xl/comments/comment1.xml','internally to xl/comments/wrong.xml'],
  ['A2 with text This is the protagonist who creates the creature.','A2 with text Wrong'],
  ["A3 with text Often mistakenly called 'Frankenstein' - that is the creator's name.",'A3 with text Wrong'],
 ])fails(await run(text=>{expect(text).toContain(from!);return text.replace(from!,to!);}));
});

test('observed comment VML results reject false external aliased duplicate missing and swapped values',async()=>{
 for(const mutate of [
  (s:CommentVmlState)=>{s.legacyId='wrong';},
  (s:CommentVmlState)=>{s.vml!.id='wrong';},
  (s:CommentVmlState)=>{s.vml!.external=true;},
  (s:CommentVmlState)=>{s.vml!.resolved=comments;},
  (s:CommentVmlState)=>{s.vml!.type=R+'/comments';},
  (s:CommentVmlState)=>{s.comments!.external=true;},
  (s:CommentVmlState)=>{s.comments!.resolved=vml;},
  (s:CommentVmlState)=>{s.comments!.id=s.vml!.id;},
  (s:CommentVmlState)=>{s.commentRows=[];},
  (s:CommentVmlState)=>{s.commentRows![0]!.text='wrong';},
  (s:CommentVmlState)=>{s.commentRows![1]!.text='wrong';},
  (s:CommentVmlState)=>{s.commentRows![1]!.ref='A2';},
  (s:CommentVmlState)=>{s.commentRows!.push({...s.commentRows![0]!});},
 ])fails(await run(text=>text,changed(read,mutate)));
});

test('wrong worksheet relationship and comment namespaces cannot satisfy the graph predicates',async()=>{
 for(const [part,mutate]of [
  [sheet,(text:string)=>text.replace(S,'urn:wrong')],
  [sheet,(text:string)=>text.replace('xmlns:r="'+R+'"','xmlns:r="urn:wrong"')],
  [sheet,(text:string)=>text.replace('r:id="anysvml"','id="anysvml"')],
  [sheet,(text:string)=>text.replace('<legacyDrawing ','<legacyDrawing xmlns="urn:wrong" ')],
  [rels,(text:string)=>text.replace('http://schemas.openxmlformats.org/package/2006/relationships','urn:wrong')],
  [comments,(text:string)=>text.replace(S,'urn:wrong')],
  [comments,(text:string)=>text.replace('<commentList>','<commentList xmlns="urn:wrong">')],
  [comments,(text:string)=>text.replace('<t>','<t xmlns="urn:wrong">')],
 ] as const){fails(await run(text=>text,changed(given,s=>{const old=s.pkg.text(part),next=mutate(old);expect(next).not.toBe(old);s.pkg.set(part,next);baseline(s);})));}
});

test('namespace aliases and rich comment text are read by expanded names without modifying their input',async()=>{
 const custom=changed(given,s=>{
  s.pkg.set(sheet,s.pkg.text(sheet).replace('xmlns:r=','xmlns:link=').replace('r:id="anysvml"','link:id="anysvml"'));
  s.pkg.set(rels,s.pkg.text(rels).replace('<Relationships xmlns=','<p:Relationships xmlns:p=').replace('</Relationships>','</p:Relationships>').replaceAll('<Relationship ','<p:Relationship '));
  s.pkg.set(comments,s.pkg.text(comments).replace('<comments xmlns=','<c:comments xmlns:c=').replace(/<(\/?)(authors|author|commentList|comment|text|t)([ >])/g,'<$1c:$2$3').replace('</comments>','</c:comments>').replace('<c:t>This is the protagonist who creates the creature.</c:t>','<c:r><c:t>This is the protagonist </c:t></c:r><c:r><c:t>who creates the creature.</c:t></c:r>'));
  baseline(s);
 });
 const result=await run(text=>text,custom);expect(result.failures).toEqual([]);expect(result.counts.cases.passed).toBe(1);
});

test('duplicate drawing IDs comment records and missing relationship dependencies fail independently of custody',async()=>{
 for(const mutate of [
  (s:CommentVmlState)=>{s.pkg.set(sheet,s.pkg.text(sheet).replace('</worksheet>','<legacyDrawing xmlns:r="'+R+'" r:id="anysvml"/></worksheet>'));},
  (s:CommentVmlState)=>{s.pkg.set(rels,s.pkg.text(rels).replace('</Relationships>','<Relationship Type="'+R+'/comments" Target="/xl/comments/comment1.xml" Id="extra"/></Relationships>'));},
  (s:CommentVmlState)=>{s.pkg.set(rels,s.pkg.text(rels).replace('Id="anysvml"','Id="comments"'));},
  (s:CommentVmlState)=>{s.pkg.set(comments,s.pkg.text(comments).replace('ref="A3"','ref="A2"'));},
  (s:CommentVmlState)=>{s.pkg.delete(vml);},
  (s:CommentVmlState)=>{s.pkg.delete(comments);},
 ])fails(await run(text=>text,changed(given,s=>{mutate(s);baseline(s);})));
});

test('read-only graph predicates detect input or opaque VML payload mutation',async()=>{
 for(const mutate of [
  (s:CommentVmlState)=>{s.input[0]=0;},
  (s:CommentVmlState)=>{s.original[0]=0;},
  (s:CommentVmlState)=>{s.pkg.set(vml,s.pkg.text(vml)+' ');},
 ])fails(await run(text=>text,changed(read,mutate)));
});
