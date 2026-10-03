import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {cropCases,cropContract,cropInput} from '../helpers/picture-crop-inputs.ts';
import {archive,xml,descendants,P,A,masked} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {tmpdir} from 'node:os';import {join} from 'node:path';
const decode=new TextDecoder();
function rects(source:string,id:number){const picture=descendants(xml(source),'pic',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===id)!;return descendants(picture,'srcRect',A);}
for(const c of cropCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await cropInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.setPictureCrop(input.shapeId,input.crop)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const previous=s.inspectPictures(),crop=s.getPictureCrop(input.shapeId);expect(crop).toEqual(c.expected.beforeCrop);crop.left=-1;expect(s.getPictureCrop(input.shapeId)).toEqual(c.expected.beforeCrop);
 expect(s.setPictureCrop(input.shapeId,input.crop)).toEqual({changed:c.expected.changed});expect(s.getPictureCrop(input.shapeId)).toEqual(c.expected.crop);expect(p.currentSlideVersion(s.partName)).toBe(version+c.expected.changed);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);
 const old=decode.decode(before.get(s.partName)),next=decode.decode(parts.get(s.partName)),oldRects=rects(old,input.shapeId),newRects=rects(next,input.shapeId);
 const remove=(source:string,nodes:{start:number;end:number}[])=>{let text=source;for(const n of [...nodes].sort((a,b)=>b.start-a.start))text=text.slice(0,n.start)+text.slice(n.end);return text;};expect(remove(next,newRects)).toBe(remove(old,oldRects));
 if(!c.expected.changed){expect(next).toBe(old);expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});}
 const records=s.inspectPictures();for(let n=0;n<previous.length;n++){const a=previous[n]!,b=records[n]!;expect({...b,crop:a.crop}).toEqual(a);}
 const dir=await mkdtemp(join(tmpdir(),'picture-crop-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);expect((await Presentation.open(bytes)).slides[0]!.getPictureCrop(input.shapeId)).toEqual(c.expected.crop);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native crop setter refuses accessors and rolls back late serialization without version changes',async()=>{
 const input=await cropInput(cropCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.setPictureCrop(input.shapeId,{...input.crop,get left(){reads++;return 0;}})).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('crop serialization failed');};try{expect(()=>s.setPictureCrop(input.shapeId,input.crop)).toThrow('crop serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.getPictureCrop(input.shapeId)).toEqual(cropCases[0].expected.beforeCrop);expect(s.setPictureCrop(input.shapeId,input.crop)).toEqual({changed:1});
});
test('native crop writes retain aliased node spelling, quotes and whitespace',async()=>{
 const c=cropCases.find((c:any)=>c.caseId==='aliased'),input=await cropInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!;
 s.setPictureCrop(input.shapeId,input.crop);const source=p.package.text(s.partName);
 expect(source).toContain("<crop:srcRect xmlns:crop=\"http://schemas.openxmlformats.org/drawingml/2006/main\" t = '20000' r='15000' b='5000' l=\"10000\"/>");
});
