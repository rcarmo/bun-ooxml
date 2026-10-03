import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {autoShapeCases,autoShapeInput} from '../helpers/autoshape-inputs.ts';
import {archive,xml,descendants,P,A} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {join} from 'node:path';import {tmpdir} from 'node:os';
const decode=new TextDecoder();
for(const c of autoShapeCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await autoShapeInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.addAutoShape(input.preset,input.geometry,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const receipt=s.addAutoShape(input.preset,input.geometry,input.options);expect(receipt).toEqual(c.expected);expect(p.currentSlideVersion(s.partName)).toBe(version+1);
 const parts=archive(p.package.toBytes());expect([...parts.keys()].sort()).toEqual([...before.keys()].sort());for(const[n,b]of before)if(n!==s.partName)expect(parts.get(n)).toEqual(b);const source=decode.decode(parts.get(s.partName)),shape=descendants(xml(source),'sp',P).find(n=>Number(descendants(n,'cNvPr',P)[0]?.attrs.id)===receipt.shapeId)!;
 expect(source.slice(0,shape.start)+source.slice(shape.end)).toBe(decode.decode(before.get(s.partName)));expect(descendants(shape,'cNvSpPr',P)[0]!.attrs).toEqual({});expect(descendants(shape,'prstGeom',A)[0]!.attrs.prst).toBe(input.preset);expect(descendants(shape,'gd',A).map(n=>n.attrs)).toEqual(Object.entries(c.expected.adjustments).map(([name,value])=>({name,fmla:'val '+value})));
 const g=input.geometry;expect(descendants(shape,'off',A)[0]!.attrs).toEqual({x:String(g.x),y:String(g.y)});expect(descendants(shape,'ext',A)[0]!.attrs).toEqual({cx:String(g.width),cy:String(g.height)});expect(descendants(shape,'t',A).map(n=>n.text).join('\n')).toBe(input.options.text??'');expect(descendants(shape,'cNvPr',P)[0]!.attrs.name).toBe(input.options.name??`AutoShape ${receipt.shapeId}`);expect(descendants(shape,'srgbClr',A).map(n=>n.attrs.val)).toEqual([input.options.fill??'F2F2F2',input.options.lineColor??'336699']);expect(descendants(shape,'ln',A)[0]!.attrs.w).toBe(String(input.options.lineWidth??12700));
 const dir=await mkdtemp(join(tmpdir(),'autoshapes-'));try{const path=join(dir,'saved.pptx');await p.save(path);const bytes=await Bun.file(path).bytes(),saved=archive(bytes);for(const[n,b]of parts)expect(saved.get(n)).toEqual(b);const reopened=await Presentation.open(bytes);expect(reopened.slides[0]!.setShapeText(receipt.shapeId,'Edited preset')).toEqual({changed:1});expect(reopened.slides[0]!.patchShapeStyle(receipt.shapeId,{fill:'FFFFFF'})).toEqual({changed:1});}finally{await rm(dir,{recursive:true,force:true});}
 receipt.geometry.x=-1;expect(p.package.text(s.partName)).toBe(source);
});
test('native AutoShape accessor refusal and late rollback preserve version',async()=>{
 const input=await autoShapeInput(autoShapeCases[0]),p=await Presentation.open(input.bytes),s=p.slides[0]!,version=p.currentSlideVersion(s.partName);let reads=0;
 expect(()=>s.addAutoShape('roundRect',input.geometry,{adjustments:{get adj(){reads++;return 1;}}})).toThrow(expect.objectContaining({code:'PPTX_AUTOSHAPE_UNSUPPORTED'}));expect(reads).toBe(0);
 const serialize=p.package.toBytes;p.package.toBytes=()=>{throw Error('AutoShape serialization failed');};try{expect(()=>s.addAutoShape(input.preset,input.geometry,input.options)).toThrow('AutoShape serialization failed');}finally{p.package.toBytes=serialize;}
 expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.addAutoShape(input.preset,input.geometry,input.options)).toEqual(autoShapeCases[0].expected);
});
