import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {deleteCases,deleteContract,deleteInput} from '../helpers/picture-delete-inputs.ts';
import {archive,xml,descendants,graph,P} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
for(const c of deleteCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await deleteInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.deletePicture(input.shapeId,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const previous=s.inspectPictures(),receipt=s.deletePicture(input.shapeId,input.options);expect(receipt).toEqual(c.expected);expect(p.currentSlideVersion(s.partName)).toBe(version+1);
 const parts=archive(p.package.toBytes()),e=c.expected;expect([...parts.keys()].sort()).toEqual([...before.keys()].filter(n=>!e.removedMedia.includes(n)).sort());
 for(const[n,b]of before)if(!deleteContract.custody.allowedChangedParts.includes(n)&&!e.removedMedia.includes(n))expect(parts.get(n)).toEqual(b);
 const source=decode.decode(before.get(s.partName)),pic=descendants(xml(source),'pic',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===input.shapeId)!;expect(decode.decode(parts.get(s.partName))).toBe(source.slice(0,pic.start)+source.slice(pic.end));
 const relPart='ppt/slides/_rels/slide1.xml.rels',rels=decode.decode(before.get(relPart));let expectedRels=rels;for(const n of xml(rels).children.filter(n=>e.removedRelationships.includes(n.attrs.Id)).sort((a,b)=>b.start-a.start))expectedRels=expectedRels.slice(0,n.start)+expectedRels.slice(n.end);expect(decode.decode(parts.get(relPart))).toBe(expectedRels);
 const g=graph(parts);for(const part of e.removedMedia)expect(g.edges.some(edge=>edge.resolved===part)).toBe(false);expect(s.inspectPictures()).toEqual(previous.filter(row=>row.shapeId!==input.shapeId));
 if(!e.removedMedia.length)expect(parts.get('[Content_Types].xml')).toEqual(before.get('[Content_Types].xml'));
 const dir=await mkdtemp(join(tmpdir(),'picture-delete-'));try{const path=join(dir,'saved.pptx');await p.save(path);const b=await Bun.file(path).bytes(),saved=archive(b);for(const[n,bytes]of parts)expect(saved.get(n)).toEqual(bytes);expect((await Presentation.open(b)).slides[0]!.inspectPictures()).toEqual(s.inspectPictures());graph(saved);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native deletion rollback restores collected media and both SVG dependency edges',async()=>{
 const input=await deleteInput(deleteCases.find((c:any)=>c.caseId==='SVG pair')),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName),serialize=p.package.toBytes;let calls=0;
 p.package.toBytes=()=>{if(++calls===5)throw Error('delete serialization failed');return serialize.call(p.package);};try{expect(()=>s.deletePicture(input.shapeId,input.options)).toThrow('delete serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.deletePicture(input.shapeId,input.options).removedMedia).toEqual(['ppt/media/image1.png','ppt/media/image1.svg']);
});
test('native deletion defaults to retention and refuses executable option accessors',async()=>{
 const input=await deleteInput(deleteCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!;let reads=0;
 expect(()=>s.deletePicture(input.shapeId,{get collectMedia(){reads++;return true;}})).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));expect(reads).toBe(0);expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});
 expect(s.deletePicture(input.shapeId)).toEqual({shapeId:input.shapeId,partName:s.partName,removedRelationships:[],removedMedia:[]});
 expect(()=>s.deletePicture(input.shapeId)).toThrow(expect.objectContaining({code:'PPTX_PICTURE_NOT_FOUND'}));
});
