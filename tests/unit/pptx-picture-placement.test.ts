import {test,expect} from 'bun:test';
import * as pptx from '../../src/pptx/index.ts';
import {placementCases,placementInput,placementContract} from '../helpers/picture-placement-inputs.ts';
import {archive,xml,descendants,graph,P,A,R} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
for(const c of placementCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await placementInput(c),p=await pptx.Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.addFittedPicture(input.payload,input.box,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const e=c.expected,calculated=pptx.calculatePicturePlacement(input.box,input.options.intrinsic,input.options.fit);expect(calculated).toEqual({geometry:e.geometry,crop:e.crop});
 const receipt=s.addFittedPicture(input.payload,input.box,input.options);expect(receipt).toEqual({shapeId:e.shapeId,partName:placementContract.slidePart,mediaPart:e.mediaPart,relationshipId:e.relationshipId,geometry:e.geometry,crop:e.crop});
 const parts=archive(p.package.toBytes()),g=graph(parts),source=decode.decode(parts.get(s.partName)),picture=descendants(xml(source),'pic',P).find(n=>descendants(n,'cNvPr',P)[0]?.attrs.id===String(e.shapeId))!;
 expect(descendants(picture,'off',A)[0]!.attrs).toEqual({x:String(e.geometry.x),y:String(e.geometry.y)});expect(descendants(picture,'ext',A)[0]!.attrs).toEqual({cx:String(e.geometry.width),cy:String(e.geometry.height)});
 const rect=descendants(picture,'srcRect',A);expect(rect).toHaveLength(Object.values(e.crop).some(v=>v!==0)?1:0);if(rect.length)expect(rect[0]!.attrs).toEqual({l:String(e.crop.left),t:String(e.crop.top),r:String(e.crop.right),b:String(e.crop.bottom)});
 expect(parts.get(e.mediaPart)).toEqual(input.source);expect(g.edges.find(r=>r.owner===s.partName&&r.id===e.relationshipId)?.type).toBe(R+'/image');expect([...parts.keys()].sort()).toEqual([...before.keys(),e.mediaPart].sort());for(const[n,b]of before)if(!placementContract.custody.allowedChangedParts.includes(n))expect(parts.get(n)).toEqual(b);
 expect(source.slice(0,picture.start)+source.slice(picture.end)).toBe(decode.decode(before.get(s.partName)));
 // Independent floating reference bounds error; serialized integer records stay exact.
 const box=input.box,intrinsic=input.options.intrinsic;
 if(input.options.fit==='contain'){const scale=Math.min(box.width/intrinsic.width,box.height/intrinsic.height);expect(Math.abs(e.geometry.width-intrinsic.width*scale)).toBeLessThanOrEqual(1);expect(Math.abs(e.geometry.height-intrinsic.height*scale)).toBeLessThanOrEqual(1);expect(e.geometry.x).toBeGreaterThanOrEqual(box.x);expect(e.geometry.y).toBeGreaterThanOrEqual(box.y);expect(e.geometry.x+e.geometry.width).toBeLessThanOrEqual(box.x+box.width);expect(e.geometry.y+e.geometry.height).toBeLessThanOrEqual(box.y+box.height);}
 if(input.options.fit==='cover'){const scale=Math.max(box.width/intrinsic.width,box.height/intrinsic.height),horizontal=(1-box.width/(intrinsic.width*scale))*50000,vertical=(1-box.height/(intrinsic.height*scale))*50000;expect(Math.abs(e.crop.left-horizontal)).toBeLessThanOrEqual(0.50001);expect(Math.abs(e.crop.top-vertical)).toBeLessThanOrEqual(0.50001);expect(e.crop.left+e.crop.right).toBeLessThan(100000);expect(e.crop.top+e.crop.bottom).toBeLessThan(100000);}
 const dir=await mkdtemp(join(tmpdir(),'picture-fit-'));try{const path=join(dir,'saved.pptx');await p.save(path);const saved=await Bun.file(path).bytes(),after=archive(saved);for(const[n,b]of parts)expect(after.get(n)).toEqual(b);const read=(await pptx.Presentation.open(saved)).slides[0]!.inspectPictures().find(p=>p.shapeId===e.shapeId)!;expect(read.transform).toEqual({...e.geometry,rotation:0,flipH:false,flipV:false});expect(read.crop).toEqual(e.crop);}finally{await rm(dir,{recursive:true,force:true});}
 receipt.geometry.width=-1;receipt.crop.left=-1;expect(s.inspectPictures().find(p=>p.shapeId===e.shapeId)!.crop).toEqual(e.crop);
});
test('native fitted insertion rolls back after full graph and placement edit',async()=>{
 const input=await placementInput(placementCases[2]),p=await pptx.Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName),serialize=p.package.toBytes;let calls=0;
 p.package.toBytes=()=>{if(++calls===3)throw Error('fit serialization failed');return serialize.call(p.package);};try{expect(()=>s.addFittedPicture(input.payload,input.box,input.options)).toThrow('fit serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.addFittedPicture(input.payload,input.box,input.options).crop).toEqual(placementCases[2].expected.crop);
});
test('native fit request refuses accessors and unknown properties without invoking caller code',async()=>{
 const input=await placementInput(placementCases[0]),p=await pptx.Presentation.open(input.bytes),s=p.slides[0]!;let reads=0;
 const accessor={contentType:'image/png',intrinsic:{width:4,height:2},get fit(){reads++;return 'contain';}};
 expect(()=>s.addFittedPicture(input.payload,input.box,accessor as typeof input.options)).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));
 expect(()=>pptx.calculatePicturePlacement(input.box,{width:4,get height(){reads++;return 2;}},'contain')).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));
 expect(()=>s.addFittedPicture(input.payload,input.box,{...input.options,alignment:'guess'})).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));expect(reads).toBe(0);expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});
});
test('native fit error distributions remain inside declared rounding bounds',()=>{
 let maxExtentError=0,maxCropError=0,count=0;
 for(const iw of [1,2,3,17,127])for(const ih of [1,2,3,11,91])for(const w of [127,999,1001])for(const h of [127,777]){
  const box={x:-10,y:20,width:w,height:h},size={width:iw,height:ih},contain=pptx.calculatePicturePlacement(box,size,'contain'),cover=pptx.calculatePicturePlacement(box,size,'cover'),scale=Math.min(w/iw,h/ih),fill=Math.max(w/iw,h/ih);
  maxExtentError=Math.max(maxExtentError,Math.abs(contain.geometry.width-iw*scale),Math.abs(contain.geometry.height-ih*scale));maxCropError=Math.max(maxCropError,Math.abs(cover.crop.left-(1-w/(iw*fill))*50000),Math.abs(cover.crop.top-(1-h/(ih*fill))*50000));count++;
 }
 expect(count).toBe(150);expect(maxExtentError).toBeLessThanOrEqual(1);expect(maxCropError).toBeLessThanOrEqual(0.50001);
});
