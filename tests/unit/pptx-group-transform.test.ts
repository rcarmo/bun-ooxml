import {test,expect} from 'bun:test';
import * as pptx from '../../src/pptx/index.ts';
import {frameCases,frameContract,frameInput} from '../helpers/group-transform-inputs.ts';
import {archive,xml,descendants,P,A,masked} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
function close(actual:{x:number;y:number},expected:{x:number;y:number}){for(const key of ['x','y'] as const)expect(Math.abs(actual[key]-expected[key])).toBeLessThanOrEqual(frameContract.tolerance.absoluteEMUs+frameContract.tolerance.relative*Math.abs(expected[key]));}
function transformSpan(source:string,id:number){const group=descendants(xml(source),'grpSp',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===id)!,props=group.children.find(n=>n.local==='grpSpPr')!,transform=props.children.find(n=>n.local==='xfrm'&&n.ns===A)!;return transform;}
for(const c of frameCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 if(c.mapping){const m=c.mapping,point=pptx.mapGroupPointChain(m.transforms,m.point);close(point,m.expected);close(pptx.unmapGroupPointChain(m.transforms,point),m.point);return;}
 const input=await frameInput(c),p=await pptx.Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.patchGroupTransform(input.shapeId,input.patch)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const previous=s.inspectPictures(),frame=s.getGroupTransform(input.shapeId);expect(frame).toEqual(c.expected.before);frame.x=-999;expect(s.getGroupTransform(input.shapeId)).toEqual(c.expected.before);expect(s.patchGroupTransform(input.shapeId,input.patch)).toEqual({changed:c.expected.changed});expect(s.getGroupTransform(input.shapeId)).toEqual(c.expected.transform);expect(p.currentSlideVersion(s.partName)).toBe(version+c.expected.changed);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);const old=decode.decode(before.get(s.partName)),next=decode.decode(parts.get(s.partName));expect(masked(next,[transformSpan(next,input.shapeId)])).toBe(masked(old,[transformSpan(old,input.shapeId)]));if(!c.expected.changed)expect(next).toBe(old);
 const rows=s.inspectPictures();for(let n=0;n<rows.length;n++){const a=previous[n]!,b=rows[n]!;expect({...b,groups:a.groups}).toEqual(a);for(const g of b.groups)if(g.shapeId===input.shapeId)expect(g.transform).toEqual(c.expected.transform);}
 const dir=await mkdtemp(join(tmpdir(),'group-frame-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);expect((await pptx.Presentation.open(bytes)).slides[0]!.getGroupTransform(input.shapeId)).toEqual(c.expected.transform);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native group transform late rollback and accessor refusal retain version',async()=>{
 const input=await frameInput(frameCases[0]),p=await pptx.Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.patchGroupTransform(input.shapeId,{get width(){reads++;return 1;}})).toThrow(expect.objectContaining({code:'PPTX_GROUP_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('group frame serialization failed');};try{expect(()=>s.patchGroupTransform(input.shapeId,input.patch)).toThrow('group frame serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.getGroupTransform(input.shapeId)).toEqual(frameCases[0].expected.before);
});
export const mappingErrors:{forward:number[];inverse:number[]}={forward:[],inverse:[]};
test('native group mapping error distribution agrees with independent polar reference',()=>{
 let count=0;
 for(const rotation of [0,5400000,10800000,16200000,1234567,9999999])for(const flipH of [false,true])for(const flipV of [false,true])for(const scale of [0.5,1,2])for(const input of [{x:-10,y:20},{x:200,y:700},{x:123.25,y:-456.5}]){
  const t={x:1000,y:-2000,width:6000*scale,height:8000*scale,childX:10,childY:20,childWidth:600,childHeight:800,rotation,flipH,flipV};
  const cx=t.x+t.width/2,cy=t.y+t.height/2,qx=t.x+(input.x-t.childX)*(t.width/t.childWidth),qy=t.y+(input.y-t.childY)*(t.height/t.childHeight),dx=(qx-cx)*(flipH?-1:1),dy=(qy-cy)*(flipV?-1:1),radius=Math.hypot(dx,dy),angle=Math.atan2(dy,dx)+rotation/60000*Math.PI/180,reference={x:cx+radius*Math.cos(angle),y:cy+radius*Math.sin(angle)};
  const actual=pptx.mapGroupPoint(t,input),inverse=pptx.unmapGroupPoint(t,actual);close(actual,reference);close(inverse,input);mappingErrors.forward.push(Math.abs(actual.x-reference.x),Math.abs(actual.y-reference.y));mappingErrors.inverse.push(Math.abs(inverse.x-input.x),Math.abs(inverse.y-input.y));count++;
 }
 expect(count).toBe(216);expect(Math.max(...mappingErrors.forward)).toBeLessThan(frameContract.tolerance.absoluteEMUs);expect(Math.max(...mappingErrors.inverse)).toBeLessThan(frameContract.tolerance.absoluteEMUs);
});
test('native group point mapping refuses singular frames, unbounded results and executable inputs',()=>{
 const frame=frameCases.find((c:any)=>c.mapping).mapping.transforms[0];let reads=0;
 for(const action of [()=>pptx.mapGroupPoint({...frame,childWidth:0},{x:1,y:1}),()=>pptx.mapGroupPoint({...frame,width:2147483647,childWidth:1},{x:2147483647,y:1}),()=>pptx.unmapGroupPoint(frame,{x:Infinity,y:0}),()=>pptx.mapGroupPoint(frame,{get x(){reads++;return 1;},y:1}),()=>pptx.mapGroupPointChain([], {x:0,y:0})])expect(action).toThrow(expect.objectContaining({code:'PPTX_GROUP_UNSUPPORTED'}));expect(reads).toBe(0);
});
