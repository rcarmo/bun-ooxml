import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {zOrderCases,zOrderInput} from '../helpers/z-order-inputs.ts';
import {archive,xml,descendants,P} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
function parent(source:string,id?:number){const d=xml(source);return id===undefined?descendants(d,'spTree',P)[0]!:descendants(d,'grpSp',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===id)!;}
const graphical=(n:{local:string})=>['sp','pic','grpSp','graphicFrame','cxnSp'].includes(n.local);
for(const c of zOrderCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await zOrderInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.reorderShapes(input.order,input.groupId)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const order=s.getShapeOrder(input.groupId);expect(order).toEqual(c.expected.before);order.length=0;expect(s.getShapeOrder(input.groupId)).toEqual(c.expected.before);expect(s.reorderShapes(input.order,input.groupId)).toEqual({changed:c.expected.changed});expect(s.getShapeOrder(input.groupId)).toEqual(c.expected.order);expect(p.currentSlideVersion(s.partName)).toBe(version+c.expected.changed);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);const old=decode.decode(before.get(s.partName)),next=decode.decode(parts.get(s.partName)),oldNodes=parent(old,input.groupId).children.filter(graphical),newNodes=parent(next,input.groupId).children.filter(graphical),oldById=new Map(oldNodes.map(n=>[Number(descendants(n,'cNvPr',P)[0]!.attrs.id),old.slice(n.start,n.end)]));
 for(const node of newNodes)expect(next.slice(node.start,node.end)).toBe(oldById.get(Number(descendants(node,'cNvPr',P)[0]!.attrs.id))!);let reconstructed=next;for(let i=newNodes.length-1;i>=0;i--){const n=newNodes[i]!,original=oldNodes[i]!;reconstructed=reconstructed.slice(0,n.start)+old.slice(original.start,original.end)+reconstructed.slice(n.end);}expect(reconstructed).toBe(old);if(!c.expected.changed)expect(next).toBe(old);
 const dir=await mkdtemp(join(tmpdir(),'z-order-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);expect((await Presentation.open(bytes)).slides[0]!.getShapeOrder(input.groupId)).toEqual(c.expected.order);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native z-order accessor refusal and late rollback preserve original slots and version',async()=>{
 const input=await zOrderInput(zOrderCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;const order=[1028,1026];Object.defineProperty(order,'0',{get(){reads++;return 1028;}});
 expect(()=>s.reorderShapes(order)).toThrow(expect.objectContaining({code:'PPTX_Z_ORDER_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('z-order serialization failed');};try{expect(()=>s.reorderShapes(input.order)).toThrow('z-order serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.getShapeOrder()).toEqual(zOrderCases[0].expected.before);
});
