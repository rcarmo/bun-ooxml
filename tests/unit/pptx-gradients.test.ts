import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {gradientCases,gradientInput} from '../helpers/gradient-inputs.ts';
import {archive,xml,descendants,P,A} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
function target(source:string){return descendants(xml(source),'sp',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===2000)!.children.find(n=>n.local==='spPr')!;}
function withoutFill(source:string){const fill=target(source).children.find(n=>['noFill','solidFill','gradFill'].includes(n.local));return fill?source.slice(0,fill.start)+source.slice(fill.end):source;}
for(const c of gradientCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await gradientInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.setLinearGradient(input.shapeId,input.gradient)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 expect(s.getLinearGradient(input.shapeId)).toEqual(c.expected.before);expect(s.setLinearGradient(input.shapeId,input.gradient)).toEqual({changed:c.expected.changed});expect(s.getLinearGradient(input.shapeId)).toEqual(c.expected.gradient);expect(p.currentSlideVersion(s.partName)).toBe(version+c.expected.changed);
 const row=s.getLinearGradient(input.shapeId)!;row.stops[0]!.color.value='changed';expect(s.getLinearGradient(input.shapeId)).toEqual(c.expected.gradient);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);const old=decode.decode(before.get(s.partName)),next=decode.decode(parts.get(s.partName));expect(withoutFill(next)).toBe(withoutFill(old));if(!c.expected.changed)expect(next).toBe(old);
 const gradient=target(next).children.find(n=>n.local==='gradFill')!,stops=descendants(gradient,'gs',A);expect(stops.map(n=>Number(n.attrs.pos))).toEqual(input.gradient.stops.map((s:any)=>s.position));for(let i=0;i<stops.length;i++){const color=stops[i]!.children[0]!,expected=input.gradient.stops[i].color;expect(color.local).toBe(expected.kind==='srgb'?'srgbClr':'schemeClr');expect(color.attrs.val).toBe(expected.value);expect(color.children.map(n=>({kind:n.local,value:Number(n.attrs.val)}))).toEqual(expected.transforms??[]);}expect(descendants(gradient,'lin',A)[0]!.attrs.ang).toBe(String(input.gradient.angle));
 const dir=await mkdtemp(join(tmpdir(),'gradients-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);expect((await Presentation.open(bytes)).slides[0]!.getLinearGradient(input.shapeId)).toEqual(c.expected.gradient);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native gradient accessor refusal and late rollback preserve version and colour references',async()=>{
 const input=await gradientInput(gradientCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.setLinearGradient(input.shapeId,{...input.gradient,get angle(){reads++;return 0;}})).toThrow(expect.objectContaining({code:'PPTX_GRADIENT_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('gradient serialization failed');};try{expect(()=>s.setLinearGradient(input.shapeId,input.gradient)).toThrow('gradient serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.getLinearGradient(input.shapeId)).toBeNull();expect(s.setLinearGradient(input.shapeId,input.gradient)).toEqual({changed:1});
});
