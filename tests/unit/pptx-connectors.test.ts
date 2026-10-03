import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {connectorCases,connectorInput} from '../helpers/connector-inputs.ts';
import {archive,xml,descendants,P,A} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
for(const c of connectorCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await connectorInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.addConnector(input.start,input.end,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const receipt=s.addConnector(input.start,input.end,input.options);expect(receipt).toEqual(c.expected);expect(p.currentSlideVersion(s.partName)).toBe(version+1);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);const source=decode.decode(parts.get(s.partName)),node=descendants(xml(source),'cxnSp',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===receipt.shapeId)!;
 expect(descendants(node,'stCxn',A)[0]!.attrs).toEqual({id:String(input.start.shapeId),idx:String(input.start.site)});expect(descendants(node,'endCxn',A)[0]!.attrs).toEqual({id:String(input.end.shapeId),idx:String(input.end.site)});
 const g=receipt.geometry;expect(descendants(node,'off',A)[0]!.attrs).toEqual({x:String(g.x),y:String(g.y)});expect(descendants(node,'ext',A)[0]!.attrs).toEqual({cx:String(g.width),cy:String(g.height)});expect(descendants(node,'xfrm',A)[0]!.attrs).toEqual({flipH:g.flipH?'1':'0',flipV:g.flipV?'1':'0'});expect(descendants(node,'prstGeom',A)[0]!.attrs.prst).toBe('line');expect(descendants(node,'srgbClr',A)[0]!.attrs.val).toBe(input.options.color);expect(descendants(node,'ln',A)[0]!.attrs.w).toBe(String(input.options.width));expect(descendants(node,'cNvPr',P)[0]!.attrs.name).toBe(input.options.name);
 expect(source.slice(0,node.start)+source.slice(node.end)).toBe(decode.decode(before.get(s.partName)));
 expect(()=>s.groupShapes([input.start.shapeId,input.end.shapeId],{x:0,y:0,width:2000,height:2000})).toThrow(expect.objectContaining({code:'PPTX_GROUP_UNSUPPORTED'}));
 const dir=await mkdtemp(join(tmpdir(),'connectors-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);const opened=await Presentation.open(bytes);expect(opened.slides[0]!.partName).toBe(s.partName);}finally{await rm(dir,{recursive:true,force:true});}
 receipt.start.x=-1;expect(p.package.text(s.partName)).toBe(source);
});
test('native attached connector accessor refusal and late rollback preserve version',async()=>{
 const input=await connectorInput(connectorCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.addConnector({get shapeId(){reads++;return 2000;},site:3},input.end,input.options)).toThrow(expect.objectContaining({code:'PPTX_CONNECTOR_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('connector serialization failed');};try{expect(()=>s.addConnector(input.start,input.end,input.options)).toThrow('connector serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.addConnector(input.start,input.end,input.options)).toEqual(connectorCases[0].expected);
});
