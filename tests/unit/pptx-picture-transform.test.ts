import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {transformCases,transformInput} from '../helpers/picture-transform-inputs.ts';
import {archive,xml,descendants,P,A,masked} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
function opening(source:string,id:number){const pic=descendants(xml(source),'pic',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===id)!,x=descendants(pic,'xfrm',A)[0]!;return {start:x.start,end:x.openEnd};}
for(const c of transformCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await transformInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.patchPictureTransform(input.shapeId,input.patch)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const previous=s.inspectPictures();expect(previous.find(r=>r.shapeId===input.shapeId)!.transform).toEqual(c.expected.before);expect(s.patchPictureTransform(input.shapeId,input.patch)).toEqual({changed:c.expected.changed});expect(p.currentSlideVersion(s.partName)).toBe(version+c.expected.changed);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);
 const old=decode.decode(before.get(s.partName)),next=decode.decode(parts.get(s.partName));expect(masked(next,[opening(next,input.shapeId)])).toBe(masked(old,[opening(old,input.shapeId)]));
 if(!c.expected.changed){expect(next).toBe(old);expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});}
 const records=s.inspectPictures();for(let n=0;n<previous.length;n++){const a=previous[n]!,b=records[n]!;expect({...b,transform:a.transform}).toEqual(a);if(b.shapeId===input.shapeId)expect(b.transform).toEqual(c.expected.transform);else expect(b).toEqual(a);}
 const dir=await mkdtemp(join(tmpdir(),'picture-orientation-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);expect((await Presentation.open(bytes)).slides[0]!.inspectPictures()).toEqual(records);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native picture transform accessor refusal and late rollback retain usable handles',async()=>{
 const input=await transformInput(transformCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.patchPictureTransform(input.shapeId,{get rotation(){reads++;return 0;}})).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('orientation serialization failed');};try{expect(()=>s.patchPictureTransform(input.shapeId,input.patch)).toThrow('orientation serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.patchPictureTransform(input.shapeId,input.patch)).toEqual({changed:1});
});
test('native picture orientation retains aliased quotes and equivalent boolean spellings',async()=>{
 const input=await transformInput(transformCases.find((c:any)=>c.caseId==='aliased')),p=await Presentation.open(input.bytes),s=p.slides[0]!;
 expect(s.patchPictureTransform(input.shapeId,{rotation:2700000,flipH:true,flipV:false})).toEqual({changed:0});expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});
 s.patchPictureTransform(input.shapeId,{rotation:5400000,flipV:true});expect(p.package.text(s.partName)).toContain("<pose:xfrm xmlns:pose=\"http://schemas.openxmlformats.org/drawingml/2006/main\" rot = '5400000' flipH='true' flipV=\"1\">");
});
