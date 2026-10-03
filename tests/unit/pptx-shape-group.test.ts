import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {groupCases,groupInput} from '../helpers/shape-group-inputs.ts';
import {archive,xml,descendants,P,A,graph} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
for(const c of groupCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await groupInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.groupShapes(input.shapeIds,input.geometry,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const prior=s.inspectPictures(),receipt=s.groupShapes(input.shapeIds,input.geometry,input.options);expect(receipt).toEqual(c.expected);expect(p.currentSlideVersion(s.partName)).toBe(version+1);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);
 const old=decode.decode(before.get(s.partName)),next=decode.decode(parts.get(s.partName)),group=descendants(xml(next),'grpSp',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===receipt.shapeId)!,identity=descendants(group,'cNvPr',P)[0]!;
 expect(identity.attrs.name).toBe(input.options.name??`Group ${receipt.shapeId}`);const transform=group.children.find(n=>n.local==='grpSpPr')!.children[0]!;expect(transform.children.map(n=>n.local)).toEqual(['off','ext','chOff','chExt']);const g=input.geometry;expect(transform.children.map(n=>n.attrs)).toEqual([{x:String(g.x),y:String(g.y)},{cx:String(g.width),cy:String(g.height)},{x:String(g.x),y:String(g.y)},{cx:String(g.width),cy:String(g.height)}]);
 const children=group.children.filter(n=>['sp','pic'].includes(n.local));expect(children.map(n=>Number(descendants(n,'cNvPr',P)[0]!.attrs.id))).toEqual(receipt.childIds);const originalTree=descendants(xml(old),'spTree',P)[0]!,selected=originalTree.children.filter(n=>input.shapeIds.includes(Number(descendants(n,'cNvPr',P)[0]?.attrs.id))),first=selected[0]!,last=selected.at(-1)!,slice=old.slice(first.start,last.end);expect(next.slice(children[0]!.start,children.at(-1)!.end)).toBe(slice);expect(next.slice(0,group.start)+slice+next.slice(group.end)).toBe(old);
 const pictures=s.inspectPictures();for(let n=0;n<prior.length;n++){const a=prior[n]!,b=pictures[n]!;expect({...b,groups:a.groups}).toEqual(a);if(receipt.childIds.includes(b.shapeId))expect(b.groups).toEqual([{shapeId:receipt.shapeId,name:identity.attrs.name!,transform:{...g,rotation:0,flipH:false,flipV:false,childX:g.x,childY:g.y,childWidth:g.width,childHeight:g.height}}]);else expect(b.groups).toEqual(a.groups);}
 graph(parts);const dir=await mkdtemp(join(tmpdir(),'shape-group-'));try{const path=join(dir,'saved.pptx');await p.save(path);const b=await Bun.file(path).bytes(),saved=archive(b);for(const[n,bytes]of parts)expect(saved.get(n)).toEqual(bytes);expect((await Presentation.open(b)).slides[0]!.inspectPictures()).toEqual(pictures);}finally{await rm(dir,{recursive:true,force:true});}
 receipt.childIds.length=0;receipt.geometry.x=999;expect(s.inspectPictures()).toEqual(pictures);
});
test('native grouping accessor refusal and late rollback preserve version and retry',async()=>{
 const input=await groupInput(groupCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.groupShapes(input.shapeIds,input.geometry,{get name(){reads++;return 'no';}})).toThrow(expect.objectContaining({code:'PPTX_GROUP_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('group serialization failed');};try{expect(()=>s.groupShapes(input.shapeIds,input.geometry,input.options)).toThrow('group serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.groupShapes(input.shapeIds,input.geometry,input.options)).toEqual(groupCases[0].expected);
});
