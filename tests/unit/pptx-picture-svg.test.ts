import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {svgCases,svgContract,svgInput} from '../helpers/picture-svg-inputs.ts';
import {archive,xml,descendants,graph,P,A,R} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
for(const c of svgCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await svgInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName),fetch=globalThis.fetch;
 globalThis.fetch=(()=>{throw Error('SVG fetched external content');}) as unknown as typeof fetch;
 try{
  if(c.errorCode){expect(()=>s.addSvgPicture(input.svg,input.fallback,input.geometry,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
  const originalSvg=input.svg.slice(),receipt=s.addSvgPicture(input.svg,input.fallback,input.geometry,input.options);expect(receipt).toEqual(c.expected);input.svg.fill(0);input.fallback.fill(0);
  const parts=archive(p.package.toBytes()),g=graph(parts),source=decode.decode(parts.get(s.partName)),picture=descendants(xml(source),'pic',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===receipt.shapeId)!,blip=descendants(picture,'blip',A)[0]!,vector=descendants(blip,'svgBlip',svgContract.svgNamespace);
  expect(vector).toHaveLength(1);expect(blip.attrs['r:embed']).toBe(receipt.relationshipId);expect(vector[0]!.attrs['r:embed']).toBe(receipt.svgRelationshipId);expect(descendants(blip,'ext',A)[0]!.attrs.uri).toBe(svgContract.extensionURI);
  expect(parts.get(receipt.mediaPart)).toEqual(input.source);expect(parts.get(receipt.svgMediaPart)).toEqual(originalSvg);expect(g.types.get(receipt.svgMediaPart)).toBe('image/svg+xml');expect(g.types.get(receipt.mediaPart)).toBe(c.request.options.contentType);
  for(const [id,part]of ([[receipt.relationshipId,receipt.mediaPart],[receipt.svgRelationshipId,receipt.svgMediaPart]] as const))expect(g.edges.find(e=>e.owner===s.partName&&e.id===id)).toEqual({owner:s.partName,id,type:R+'/image',target:'../media/'+part.split('/').at(-1),resolved:part,external:false});
  expect([...parts.keys()].sort()).toEqual([...before.keys(),receipt.mediaPart,receipt.svgMediaPart].sort());for(const[n,b]of before)if(!svgContract.custody.allowedChangedParts.includes(n))expect(parts.get(n)).toEqual(b);expect(source.slice(0,picture.start)+source.slice(picture.end)).toBe(decode.decode(before.get(s.partName)));
  const relPart='ppt/slides/_rels/slide1.xml.rels',rels=decode.decode(parts.get(relPart));let restored=rels;for(const n of xml(rels).children.filter(n=>[receipt.relationshipId,receipt.svgRelationshipId].includes(n.attrs.Id!)).sort((a,b)=>b.start-a.start))restored=restored.slice(0,n.start)+restored.slice(n.end);expect(restored).toBe(decode.decode(before.get(relPart)));
  expect(()=>s.replacePicture(receipt.shapeId,input.source,{contentType:c.request.options.contentType})).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));
  const dir=await mkdtemp(join(tmpdir(),'picture-svg-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);expect((await Presentation.open(bytes)).slides[0]!.inspectPictures()).toEqual(s.inspectPictures());}finally{await rm(dir,{recursive:true,force:true});}
 }finally{globalThis.fetch=fetch;}
});
test('native SVG insertion rollback removes both dependency edges and payloads',async()=>{
 const input=await svgInput(svgCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName),serialize=p.package.toBytes;let calls=0;
 p.package.toBytes=()=>{if(++calls===5)throw Error('SVG pair serialization failed');return serialize.call(p.package);};try{expect(()=>s.addSvgPicture(input.svg,input.fallback,input.geometry,input.options)).toThrow('SVG pair serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.addSvgPicture(input.svg,input.fallback,input.geometry,input.options)).toEqual(svgCases[0].expected);
});
test('native SVG bounded bytes and escaped paint refusal preserve package state',async()=>{
 const input=await svgInput(svgCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!;
 for(const bytes of [new Uint8Array(),new Uint8Array(1024*1024+1),new Uint8Array([255]),new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><rect fill="u\\72l(https://invalid.example)"/></svg>')])expect(()=>s.addSvgPicture(bytes,input.fallback,input.geometry,input.options)).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});
});
