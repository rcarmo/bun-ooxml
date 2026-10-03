import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {opacityCases,opacityInput} from '../helpers/opacity-inputs.ts';
import {archive,xml,descendants,P,A} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
function alpha(source:string,id:number,profile:string){const target=descendants(xml(source),profile==='shape'?'sp':'pic',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===id)!;if(profile==='picture')return descendants(target,'alphaModFix',A);const props=target.children.find(n=>n.local==='spPr')!,fill=props.children.find(n=>n.local==='solidFill')!;return descendants(fill,'alpha',A);}
function strip(source:string,id:number,profile:string){let s=source;for(const n of alpha(source,id,profile).sort((a,b)=>b.start-a.start))s=s.slice(0,n.start)+s.slice(n.end);return s;}
for(const c of opacityCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await opacityInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName),read=()=>input.profile==='shape'?s.getShapeOpacity(input.shapeId):s.getPictureTransparency(input.shapeId),write=()=>input.profile==='shape'?s.setShapeOpacity(input.shapeId,input.value):s.setPictureTransparency(input.shapeId,input.value);
 if(c.errorCode){expect(write).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 expect(read()).toBe(c.expected.before);expect(write()).toEqual({changed:c.expected.changed});expect(read()).toBe(c.expected.value);expect(p.currentSlideVersion(s.partName)).toBe(version+c.expected.changed);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);const old=decode.decode(before.get(s.partName)),next=decode.decode(parts.get(s.partName));expect(strip(next,input.shapeId,input.profile)).toBe(strip(old,input.shapeId,input.profile));if(!c.expected.changed)expect(next).toBe(old);const node=alpha(next,input.shapeId,input.profile)[0];if(node)expect(Number(node.attrs[input.profile==='shape'?'val':'amt'])).toBe(input.profile==='shape'?input.value:100000-input.value);
 const dir=await mkdtemp(join(tmpdir(),'opacity-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);const reopened=(await Presentation.open(bytes)).slides[0]!;expect(input.profile==='shape'?reopened.getShapeOpacity(input.shapeId):reopened.getPictureTransparency(input.shapeId)).toBe(c.expected.value);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native opacity late rollback preserves direct references and version',async()=>{
 const input=await opacityInput(opacityCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName),serialize=p.package.toBytes;
 p.package.toBytes=()=>{throw Error('opacity serialization failed');};try{expect(()=>s.setShapeOpacity(input.shapeId,input.value)).toThrow('opacity serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.getShapeOpacity(input.shapeId)).toBe(100000);
});
test('native opacity expands an empty RGB leaf and rolls back picture effects independently',async()=>{
 const p=Presentation.create(),s=p.addTextSlide('Native'),shape=s.addAutoShape('rect',{x:0,y:0,width:1000,height:1000});
 expect(s.setShapeOpacity(shape.shapeId,40000)).toEqual({changed:1});expect(s.getShapeOpacity(shape.shapeId)).toBe(40000);expect(p.package.text(s.partName)).toContain('<a:srgbClr val="F2F2F2"><a:alpha xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" val="40000"/></a:srgbClr>');
 const input=await opacityInput(opacityCases.find((c:any)=>c.caseId==='existing picture')),deck=await Presentation.open(input.bytes),slide=deck.slides[0]!,version=deck.currentSlideVersion(slide.partName),serialize=deck.package.toBytes;
 deck.package.toBytes=()=>{throw Error('picture transparency serialization failed');};try{expect(()=>slide.setPictureTransparency(input.shapeId,input.value)).toThrow('picture transparency serialization failed');}finally{deck.package.toBytes=serialize;}
 expect(deck.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(deck.currentSlideVersion(slide.partName)).toBe(version);expect(slide.getPictureTransparency(input.shapeId)).toBe(65000);
});
