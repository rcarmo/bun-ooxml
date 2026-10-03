import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {outlineCases,outlineInput} from '../helpers/outline-inputs.ts';
import {archive,xml,descendants,P,A,masked} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
function line(source:string,id:number){const n=[...descendants(xml(source),'sp',P),...descendants(xml(source),'cxnSp',P)].find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===id)!;return n.children.find(c=>c.local==='spPr')!.children.find(c=>c.local==='ln'&&c.ns===A)!;}
function without(source:string,id:number){const ln=line(source,id);return masked(source,[{start:ln.start,end:ln.openEnd},...ln.children.filter(n=>['round','bevel','miter','headEnd','tailEnd'].includes(n.local))]);}
for(const c of outlineCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await outlineInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.patchOutlineStyle(input.shapeId,input.patch)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 expect(s.getOutlineStyle(input.shapeId)).toEqual(c.expected.before);expect(s.patchOutlineStyle(input.shapeId,input.patch)).toEqual({changed:c.expected.changed});expect(s.getOutlineStyle(input.shapeId)).toEqual(c.expected.outline);expect(p.currentSlideVersion(s.partName)).toBe(version+c.expected.changed);const read=s.getOutlineStyle(input.shapeId);if(read.headEnd)read.headEnd.type='none';expect(s.getOutlineStyle(input.shapeId)).toEqual(c.expected.outline);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);const old=decode.decode(before.get(s.partName)),next=decode.decode(parts.get(s.partName));
 const remove=(source:string)=>{let value=source;for(const n of line(source,input.shapeId).children.filter(n=>['round','bevel','miter','headEnd','tailEnd'].includes(n.local)).sort((a,b)=>b.start-a.start))value=value.slice(0,n.start)+value.slice(n.end);const ln=line(value,input.shapeId);return value.slice(0,ln.start)+'<LINE>'+value.slice(ln.openEnd);};expect(remove(next)).toBe(remove(old));expect(line(next,input.shapeId).attrs.w).toBe(line(old,input.shapeId).attrs.w);if(!c.expected.changed)expect(next).toBe(old);
 if(c.caseId==='aliased scalar update')expect(next).toContain("<outline:headEnd xmlns:outline=\"http://schemas.openxmlformats.org/drawingml/2006/main\" type = 'diamond' w='lg' len='med'/>");
 const ln=line(next,input.shapeId),e=c.expected.outline;expect(ln.attrs.cap??'flat').toBe(e.cap);expect(ln.attrs.cmpd??'sng').toBe(e.compound);if(e.join){const n=ln.children.find(n=>n.local===e.join.type)!;expect(n).toBeDefined();if(e.join.type==='miter')expect(n.attrs.lim).toBe(String(e.join.limit));}for(const key of ['headEnd','tailEnd'] as const){const n=ln.children.find(n=>n.local===key);if(e[key])expect({type:n!.attrs.type,w:n!.attrs.w,len:n!.attrs.len}).toEqual({type:e[key].type,w:e[key].width,len:e[key].length});else expect(n).toBeUndefined();}
 const dir=await mkdtemp(join(tmpdir(),'outlines-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);expect((await Presentation.open(bytes)).slides[0]!.getOutlineStyle(input.shapeId)).toEqual(c.expected.outline);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native outline accessor refusal and late rollback preserve version',async()=>{
 const input=await outlineInput(outlineCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.patchOutlineStyle(input.shapeId,{get cap(){reads++;return 'rnd' as const;}})).toThrow(expect.objectContaining({code:'PPTX_OUTLINE_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('outline serialization failed');};try{expect(()=>s.patchOutlineStyle(input.shapeId,input.patch)).toThrow('outline serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.getOutlineStyle(input.shapeId)).toEqual(outlineCases[0].expected.before);
});
