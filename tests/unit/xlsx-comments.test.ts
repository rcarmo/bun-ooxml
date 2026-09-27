import {test,expect} from 'bun:test';
import * as api from '../../src/index.ts';
import {fixturePath} from '../../scripts/fixture-inputs.ts';
import {addPart,addRelationship} from '../../src/opc/index.ts';
const {OpcPackage,Workbook}=api;
const inspect=(pkg:api.OpcPackage,part=sheet)=>api.inspectWorksheetComments(pkg,part);
const id='fixture-264be55e012d4bc2b3bf25e59824fdd30022e94f70869ea7d6ad960b803a902f',sheet='xl/worksheets/sheet1.xml',rels='xl/worksheets/_rels/sheet1.xml.rels',comments='xl/comments/comment1.xml',vml='xl/drawings/commentsDrawing1.vml';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',C='application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml';
const fixture=()=>OpcPackage.open(fixturePath(id));
const textA="This is the protagonist who creates the creature.",textB="Often mistakenly called 'Frankenstein' - that is the creator's name.";

test('native comment inspection returns detached exact authors refs text and separate VML relationship without writes',async()=>{
 const p=await fixture(),before=p.toBytes(),out=inspect(p);expect(out).toEqual({worksheetPart:sheet,commentsRelationship:{id:'comments',type:R+'/comments',part:comments},vmlRelationship:{id:'anysvml',type:R+'/vmlDrawing',part:vml},comments:[{reference:'A2',authorId:0,author:'Test Author',text:textA},{reference:'A3',authorId:0,author:'Test Author',text:textB}]});expect(p.toBytes()).toEqual(before);
 out.comments[0]!.text='wrong';out.comments.push({reference:'B1',authorId:0,author:'Fake',text:'fake'});out.commentsRelationship!.part='wrong';expect(inspect(p).comments.map(c=>c.text)).toEqual([textA,textB]);expect(inspect(p).commentsRelationship!.part).toBe(comments);expect(p.toBytes()).toEqual(before);
 const after=await OpcPackage.open(p.toBytes());expect(inspect(after)).toEqual(inspect(p));
});

test('empty comment graphs and comments without VML remain distinct from malformed dependencies',()=>{
 const p=Workbook.create().package,before=p.toBytes();expect(inspect(p)).toEqual({worksheetPart:sheet,commentsRelationship:null,vmlRelationship:null,comments:[]});expect(p.toBytes()).toEqual(before);
 addPart(p,comments,`<comments xmlns="${S}"><authors><author>A</author></authors><commentList><comment ref="XFD1048576" authorId="0"><text><t>edge</t></text></comment></commentList></comments>`,C);addRelationship(p,sheet,R+'/comments','../comments/comment1.xml');const out=inspect(p);expect(out.comments).toEqual([{reference:'XFD1048576',authorId:0,author:'A',text:'edge'}]);expect(out.vmlRelationship).toBeNull();expect(out.commentsRelationship!.part).toBe(comments);
});

test('rich text alias and UTF16 comment parts read in document order without normalising original bytes',async()=>{
 const p=await fixture();p.set(sheet,p.text(sheet).replace('xmlns:r=','xmlns:link=').replace('r:id="anysvml"','link:id="anysvml"'));
 let xml=`<c:comments xmlns:c="${S}"><c:authors><c:author>  雪 &amp; Co  </c:author></c:authors><c:commentList><c:comment ref="B9" authorId="0"><c:text><c:r><c:rPr><c:b/><c:color rgb="FFFF0000"/></c:rPr><c:t xml:space="preserve"> A&amp; </c:t></c:r><c:r><c:t>😀</c:t></c:r></c:text></c:comment><c:comment ref="A1" authorId="0"><c:text><c:t/></c:text></c:comment></c:commentList></c:comments>`;
 const {utf16}=await import('../fixtures/admission.ts');p.set(comments,utf16('<?xml version="1.0" encoding="UTF-16"?>'+xml));const before=p.toBytes(),out=inspect(p);expect(out.comments).toEqual([{reference:'B9',authorId:0,author:'  雪 & Co  ',text:' A& 😀'},{reference:'A1',authorId:0,author:'  雪 & Co  ',text:''}]);expect(out.vmlRelationship!.id).toBe('anysvml');expect(p.toBytes()).toEqual(before);
});

test('external ambiguous mismatched and dangling comment or VML dependencies refuse unchanged',async()=>{
 for(const mutate of [
  (p:api.OpcPackage)=>p.set(rels,p.text(rels).replace('Id="comments"','Id="comments" TargetMode="External"')),
  (p:api.OpcPackage)=>p.set(rels,p.text(rels).replace('Id="anysvml"','Id="anysvml" TargetMode="External"')),
  (p:api.OpcPackage)=>addRelationship(p,sheet,R+'/comments','../comments/comment1.xml',{id:'second'}),
  (p:api.OpcPackage)=>addRelationship(p,sheet,R+'/vmlDrawing','../drawings/commentsDrawing1.vml',{id:'second'}),
  (p:api.OpcPackage)=>p.set(sheet,p.text(sheet).replace('r:id="anysvml"','r:id="missing"')),
  (p:api.OpcPackage)=>p.set(sheet,p.text(sheet).replace(/<legacyDrawing[^>]*\/>/,'')),
  (p:api.OpcPackage)=>p.set(sheet,p.text(sheet).replace('</worksheet>',`<legacyDrawing xmlns:r="${R}" r:id="anysvml"/></worksheet>`)),
  (p:api.OpcPackage)=>p.set(rels,p.text(rels).replace('/vmlDrawing','/drawing')),
  (p:api.OpcPackage)=>p.delete(vml),
  (p:api.OpcPackage)=>p.delete(comments),
 ]){const p=await fixture();mutate(p);const before=p.names().map(n=>[n,p.get(n)]);expect(()=>inspect(p)).toThrow();expect(p.names().map(n=>[n,p.get(n)])).toEqual(before);}
});

test('wrong expanded namespaces and part content types never fall back to local-name inspection',async()=>{
 for(const mutate of [
  (p:api.OpcPackage)=>p.set(sheet,p.text(sheet).replace(S,'urn:wrong')),
  (p:api.OpcPackage)=>p.set(sheet,p.text(sheet).replace('xmlns:r="'+R+'"','xmlns:r="urn:wrong"')),
  (p:api.OpcPackage)=>p.set(sheet,p.text(sheet).replace('r:id="anysvml"','id="anysvml"')),
  (p:api.OpcPackage)=>p.set(comments,p.text(comments).replace(S,'urn:wrong')),
  (p:api.OpcPackage)=>p.set(comments,p.text(comments).replace('ref="A2"','xmlns:q="urn:wrong" q:ref="A2"')),
  (p:api.OpcPackage)=>p.set('[Content_Types].xml',p.text('[Content_Types].xml').replace(C,'application/xml')),
  (p:api.OpcPackage)=>p.set('[Content_Types].xml',p.text('[Content_Types].xml').replace('application/vnd.openxmlformats-officedocument.vmlDrawing','application/xml')),
 ]){const p=await fixture();mutate(p);const before=p.toBytes();expect(()=>inspect(p)).toThrow();expect(p.toBytes()).toEqual(before);}
});

test('invalid duplicate and non-cell comment references or author indices refuse without partial lists',async()=>{
 for(const [from,to]of [
  ['ref="A3"','ref="A2"'],...['A0','XFE1','A1048577','A1:B2','$A$1','a2','A02','Sheet1!A1',' A2 '].map(ref=>['ref="A2"',`ref="${ref}"`]),
  ...['-1','1','NaN','00','0.5','9007199254740992'].map(id=>['authorId="0"',`authorId="${id}"`]),
 ]){const p=await fixture();p.set(comments,p.text(comments).replace(from!,to!));const before=p.toBytes();expect(()=>inspect(p)).toThrow();expect(p.toBytes()).toEqual(before);}
});

test('unknown comment content phonetics nested text and malformed rich runs refuse instead of returning partial text',async()=>{
 for(const [from,to]of [
  ['</comments>','<extLst/></comments>'],['<authors>','<authors><unknown/>'],['<commentList>','<commentList/><commentList>'],['<text><t>','<text><rPh><t>hidden</t></rPh><t>'],['<text><t>','<text><unknown/> <t>'],['<t>This','<t><b/>This'],['<text><t>','<text><t/> <t>'],['<text><t>','<text><r><unknown/></r><t>'],['<text><t>','<text><!--hidden--><t>'],
 ]){const p=await fixture();p.set(comments,p.text(comments).replace(from!,to!));const before=p.toBytes();expect(()=>inspect(p)).toThrow();expect(p.toBytes()).toEqual(before);}
});

test('fresh inspections observe live metadata changes and protection does not turn reads into mutations',async()=>{
 const p=await fixture(),first=inspect(p);p.set(comments,p.text(comments).replace(textA,'changed'));p.set(sheet,p.text(sheet).replace('</worksheet>','<sheetProtection sheet="1"/></worksheet>'));const before=p.toBytes();expect(inspect(p).comments[0]!.text).toBe('changed');expect(first.comments[0]!.text).toBe(textA);expect(p.toBytes()).toEqual(before);
 p.set(vml,'opaque non-XML drawing payload');const opaque=p.toBytes();expect(inspect(p).vmlRelationship!.part).toBe(vml);expect(p.toBytes()).toEqual(opaque);
});

test('bounded worksheet-comment inputs refuse oversized XML comment counts and invalid part arguments',async()=>{
 const p=await fixture();for(const part of ['', '../xl/worksheets/sheet1.xml','/xl/worksheets/sheet1.xml','word/document.xml',null])expect(()=>inspect(p,part as string)).toThrow();
 p.set(comments,`<comments xmlns="${S}"><authors><author>A</author></authors><commentList>${Array.from({length:10001},(_,i)=>`<comment ref="A${i+1}" authorId="0"><text><t>x</t></text></comment>`).join('')}</commentList></comments>`);const before=p.toBytes();expect(()=>inspect(p)).toThrow(expect.objectContaining({code:'xlsx-comments-unsupported'}));expect(p.toBytes()).toEqual(before);
 p.set(comments,`<comments xmlns="${S}">`+' '.repeat(8*1024*1024)+'</comments>');const large=p.toBytes();expect(()=>inspect(p)).toThrow();expect(p.toBytes()).toEqual(large);
});

test('VML-only and empty comment parts report their distinct graphs without manufacturing comments',async()=>{
 const p=await fixture();p.set(rels,p.text(rels).replace(/<Relationship [^>]*Id="comments"[^>]*\/>/,''));const before=p.toBytes();expect(inspect(p).comments).toEqual([]);expect(inspect(p).commentsRelationship).toBeNull();expect(inspect(p).vmlRelationship!.part).toBe(vml);expect(p.toBytes()).toEqual(before);
 const q=await fixture();q.set(comments,`<comments xmlns="${S}"><authors/><commentList/></comments>`);const empty=q.toBytes();expect(inspect(q).comments).toEqual([]);expect(inspect(q).commentsRelationship!.part).toBe(comments);expect(q.toBytes()).toEqual(empty);
});

test('relationship lookalike attributes and fragment targets refuse without changing part payloads',async()=>{
 for(const [from,to]of [
  ['Id="anysvml"','xmlns:q="urn:q" q:Id="anysvml"'],
  ['Id="anysvml"','Id="anysvml" xmlns:q="urn:q" q:Id="extra"'],
  ['Target="/xl/comments/comment1.xml"','Target="/xl/comments/comment1.xml#fragment"'],
  ['Target="/xl/drawings/commentsDrawing1.vml"','Target="/xl/drawings/commentsDrawing1.vml?query"'],
 ]){const p=await fixture();p.set(rels,p.text(rels).replace(from!,to!));const before=p.names().map(n=>[n,p.get(n)]);expect(()=>inspect(p)).toThrow();expect(p.names().map(n=>[n,p.get(n)])).toEqual(before);}
});

test('refused metadata can be corrected for a later inspection without changing earlier detached results',async()=>{
 const p=await fixture(),original=p.text(comments),first=inspect(p);p.set(comments,original.replace('authorId="0"','authorId="9"'));const invalid=p.toBytes();expect(()=>inspect(p)).toThrow(expect.objectContaining({code:'xlsx-comments-unsupported'}));expect(p.toBytes()).toEqual(invalid);p.set(comments,original.replace(textA,'Corrected'));expect(inspect(p).comments[0]!.text).toBe('Corrected');expect(first.comments[0]!.text).toBe(textA);
});
