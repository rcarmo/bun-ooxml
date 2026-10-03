import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {insertionCases,insertionInput,insertionContract} from '../helpers/picture-insertion-inputs.ts';
import {archive,xml,descendants,P,A,R,graph} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';
const decode=new TextDecoder();
function unchanged(before:Map<string,Uint8Array>,after:Map<string,Uint8Array>,except:string[]=[]){for(const[p,b]of before)if(!except.includes(p))expect(after.get(p)).toEqual(b);}
for(const c of insertionCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await insertionInput(c),p=await Presentation.open(input.bytes),slide=p.slides[0]!,before=archive(input.bytes);
 if(c.errorCode){
  expect(()=>slide.addPicture(input.payload,input.geometry,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));
  const after=archive(p.package.toBytes());expect([...after.keys()].sort()).toEqual([...before.keys()].sort());unchanged(before,after);expect(slide.index).toBe(0);expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});return;
 }
 for(const expected of c.expected){
  const receipt=slide.addPicture(input.payload,input.geometry,input.options);
  expect(receipt).toEqual({shapeId:expected.shapeId,partName:expected.slidePart,mediaPart:expected.mediaPart,relationshipId:expected.relationshipId});
  input.payload.fill(0); // Defensive copy; repeat uses the same original payload again.
  const parts=archive(p.package.toBytes()),g=graph(parts),source=decode.decode(parts.get(expected.slidePart)),picture=descendants(xml(source),'pic',P).find(n=>descendants(n,'cNvPr',P)[0]?.attrs.id===String(expected.shapeId))!;
  expect(picture).toBeDefined();const id=descendants(picture,'cNvPr',P)[0]!;expect(id.attrs.name).toBe(expected.name);expect(id.attrs.descr).toBe(expected.description);
  const off=descendants(picture,'off',A)[0]!,ext=descendants(picture,'ext',A)[0]!;expect(off.attrs).toEqual({x:String(expected.geometry.x),y:String(expected.geometry.y)});expect(ext.attrs).toEqual({cx:String(expected.geometry.width),cy:String(expected.geometry.height)});
  expect(g.edges.find(e=>e.owner===expected.slidePart&&e.id===expected.relationshipId)).toEqual({owner:expected.slidePart,id:expected.relationshipId,type:R+'/image',target:expected.target,resolved:expected.mediaPart,external:false});
  expect(g.types.get(expected.mediaPart)).toBe(expected.contentType);expect(parts.get(expected.mediaPart)).toEqual(input.source);
  input.payload=input.source.slice();
 }
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys(),...c.expected.map((e:any)=>e.mediaPart)].sort());unchanged(before,parts,insertionContract.custody.allowedChangedParts);
 const original=decode.decode(before.get(insertionContract.slidePart)),changed=decode.decode(parts.get(insertionContract.slidePart)),tree=xml(changed);
 const added=descendants(tree,'pic',P).filter(n=>c.expected.some((e:any)=>String(e.shapeId)===descendants(n,'cNvPr',P)[0]?.attrs.id));
 // Remove exactly the appended nodes and compare the complete original slide XML.
 let recovered=changed;for(const n of [...added].sort((a,b)=>b.start-a.start))recovered=recovered.slice(0,n.start)+recovered.slice(n.end);expect(recovered).toBe(original);
 const relPart='ppt/slides/_rels/slide1.xml.rels',oldEdges=graph(before).edges.filter(e=>e.owner===insertionContract.slidePart),newEdges=graph(parts).edges.filter(e=>e.owner===insertionContract.slidePart);for(const edge of oldEdges)expect(newEdges).toContainEqual(edge);
 const newRels=decode.decode(parts.get(relPart)),newRelNodes=xml(newRels).children.filter(n=>c.expected.some((e:any)=>e.relationshipId===n.attrs.Id));let restoredRels=newRels;for(const n of [...newRelNodes].sort((a,b)=>b.start-a.start))restoredRels=restoredRels.slice(0,n.start)+restoredRels.slice(n.end);expect(restoredRels).toBe(decode.decode(before.get(relPart)));
 if(c.caseId==='terminal-extension'){const shapeTree=descendants(tree,'spTree',P)[0]!;expect(shapeTree.children.at(-1)!.local).toBe('extLst');expect(added.at(-1)!.end).toBeLessThanOrEqual(shapeTree.children.at(-1)!.start);}
 if(c.caseId==='conflicting-MIME')expect(decode.decode(parts.get('[Content_Types].xml'))).toContain('Extension="png" ContentType="application/octet-stream"');
 const dir=await mkdtemp(join(tmpdir(),'pptx-picture-add-'));try{const path=join(dir,'saved.pptx');await p.save(path);const saved=await Bun.file(path).bytes(),savedParts=archive(saved),reopened=await Presentation.open(saved);unchanged(parts,savedParts);expect([...savedParts.keys()].sort()).toEqual([...parts.keys()].sort());expect(reopened.slides[0]!.inspectPictures()).toEqual(slide.inspectPictures());}finally{await rm(dir,{recursive:true,force:true});}
});
test('native picture insertion rolls back serialization failure and retains handle version',async()=>{
 const input=await insertionInput(insertionCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(p.package.toBytes()),version=p.currentSlideVersion(s.partName),serialize=p.package.toBytes;let calls=0;
 // addPart and addRelationship serialize first; fail only after the slide is edited.
 p.package.toBytes=()=>{if(++calls===3)throw Error('injected picture serialization');return serialize.call(p.package);};try{expect(()=>s.addPicture(input.payload,input.geometry,input.options)).toThrow('injected picture serialization');}finally{p.package.toBytes=serialize;}
 unchanged(before,archive(p.package.toBytes()));expect(p.currentSlideVersion(s.partName)).toBe(version);expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});
 expect(s.addPicture(input.payload,input.geometry,input.options).shapeId).toBe(insertionCases[0].expected[0].shapeId);
});
test('native picture insertion refuses non-byte values, accessor options and oversized payloads before editing',async()=>{
 const input=await insertionInput(insertionCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!;let getterCalls=0;
 const accessor={contentType:'image/png',get name(){getterCalls++;return 'side effect';}};
 for(const action of [()=>s.addPicture('path.png' as unknown as Uint8Array,input.geometry,input.options),()=>s.addPicture(new Uint8Array(64*1024*1024+1),input.geometry,input.options),()=>s.addPicture(input.payload,input.geometry,accessor as typeof input.options)])expect(action).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));
 expect(getterCalls).toBe(0);expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});
});
test('native picture insertion creates missing slide relationships and uses detached default metadata',async()=>{
 const input=await insertionInput(insertionCases[0]),p=Presentation.create(),s=p.addTextSlide('Native slide'),before=archive(p.package.toBytes());
 const receipt=s.addPicture(input.payload,input.geometry,{contentType:'image/png'}),saved=archive(p.package.toBytes()),picture=s.inspectPictures()[0]!;
 expect(picture.shapeId).toBe(receipt.shapeId);expect(picture.name).toBe(`Picture ${receipt.shapeId}`);expect(picture.description).toBeNull();expect(picture.embedded!.partName).toBe(receipt.mediaPart);expect(saved.get(receipt.mediaPart)).toEqual(input.source);unchanged(before,saved,[s.partName,'[Content_Types].xml',s.partName.replace('/slides/','/slides/_rels/')+'.rels']);
 receipt.shapeId=-1;expect(s.inspectPictures()[0]!.shapeId).toBe(picture.shapeId);
 expect((await Presentation.open(p.package.toBytes())).slides[0]!.inspectPictures()).toEqual(s.inspectPictures());
});
