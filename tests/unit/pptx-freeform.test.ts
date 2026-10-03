import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {freeformCases,freeformInput} from '../helpers/freeform-inputs.ts';
import {archive,xml,descendants,P,A} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
for(const c of freeformCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await freeformInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.addFreeform(input.geometry,input.path,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const receipt=s.addFreeform(input.geometry,input.path,input.options);expect(receipt).toEqual(c.expected);expect(p.currentSlideVersion(s.partName)).toBe(version+1);const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);
 const source=decode.decode(parts.get(s.partName)),shape=descendants(xml(source),'sp',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===receipt.shapeId)!,custom=descendants(shape,'custGeom',A);expect(custom).toHaveLength(1);expect(descendants(shape,'prstGeom',A)).toHaveLength(0);expect(descendants(shape,'cNvSpPr',P)[0]!.attrs).toEqual({});expect(source.slice(0,shape.start)+source.slice(shape.end)).toBe(decode.decode(before.get(s.partName)));
 const path=descendants(shape,'path',A)[0]!;expect(path.attrs).toEqual({w:String(input.path.width),h:String(input.path.height),fill:(input.options.fill??'none')==='none'?'none':'norm',stroke:'1',extrusionOk:'0'});expect(path.children.map(n=>n.local)).toEqual(input.path.commands.map((c:any)=>({move:'moveTo',line:'lnTo',close:'close'}[c.op as 'move'|'line'|'close'])));expect(path.children.map(n=>n.local==='close'?{op:'close'}:{op:n.local==='moveTo'?'move':'line',x:Number(n.children[0]!.attrs.x),y:Number(n.children[0]!.attrs.y)})).toEqual(input.path.commands);
 const g=input.geometry;expect(descendants(shape,'off',A)[0]!.attrs).toEqual({x:String(g.x),y:String(g.y)});expect(descendants(shape,'ext',A)[0]!.attrs).toEqual({cx:String(g.width),cy:String(g.height)});expect(descendants(shape,'cNvPr',P)[0]!.attrs.name).toBe(input.options.name??`Freeform ${receipt.shapeId}`);expect(descendants(shape,'srgbClr',A).map(n=>n.attrs.val)).toEqual((input.options.fill??'none')==='none'?[input.options.lineColor??'336699']:[input.options.fill,input.options.lineColor??'336699']);
 const dir=await mkdtemp(join(tmpdir(),'freeform-'));try{const output=join(dir,'saved.pptx');await p.save(output);const bytes=await Bun.file(output).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);expect((await Presentation.open(bytes)).package.text(s.partName)).toBe(source);}finally{await rm(dir,{recursive:true,force:true});}
 receipt.path.commands.length=0;receipt.geometry.x=-1;expect(p.package.text(s.partName)).toBe(source);
});
test('native freeform accessor refusal and late rollback preserve version',async()=>{
 const input=await freeformInput(freeformCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.addFreeform(input.geometry,{width:100,height:100,commands:[{op:'move',get x(){reads++;return 0;},y:0},{op:'line',x:1,y:1}]},{fill:'none'})).toThrow(expect.objectContaining({code:'PPTX_FREEFORM_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('freeform serialization failed');};try{expect(()=>s.addFreeform(input.geometry,input.path,input.options)).toThrow('freeform serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.addFreeform(input.geometry,input.path,input.options)).toEqual(freeformCases[0].expected);
});
