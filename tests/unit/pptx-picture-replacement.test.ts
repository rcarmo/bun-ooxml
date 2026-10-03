import {test,expect} from 'bun:test';
import {Presentation} from '../../src/pptx/index.ts';
import {replacementCases,replacementContract,replacementInput} from '../helpers/picture-replacement-inputs.ts';
import {archive,xml,descendants,graph,P,A,R} from '../helpers/contract20-oracle.ts';
import {mkdtemp,rm} from 'node:fs/promises';import {tmpdir} from 'node:os';import {join} from 'node:path';
const decode=new TextDecoder();
function custody(before:Map<string,Uint8Array>,after:Map<string,Uint8Array>,except:string[]=[]){for(const[p,b]of before)if(!except.includes(p))expect(after.get(p)).toEqual(b);}
function selected(parts:Map<string,Uint8Array>,id:number){const source=decode.decode(parts.get(replacementContract.slidePart)),picture=descendants(xml(source),'pic',P).find(n=>descendants(n,'cNvPr',P)[0]?.attrs.id===String(id))!;return {source,picture,blip:descendants(picture,'blip',A)[0]!};}
for(const c of replacementCases)test(`${c.scenarioId} [${c.caseId}]`,async()=>{
 const input=await replacementInput(c),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName);
 if(c.errorCode){expect(()=>s.replacePicture(input.shapeId,input.payload,input.options)).toThrow(expect.objectContaining({code:c.errorCode}));const after=archive(p.package.toBytes());expect([...after.keys()].sort()).toEqual([...before.keys()].sort());custody(before,after);expect(p.currentSlideVersion(s.partName)).toBe(version);return;}
 const previous=s.inspectPictures(),receipt=s.replacePicture(input.shapeId,input.payload,input.options),expected=c.expected;
 expect(receipt).toEqual({shapeId:expected.shapeId,partName:expected.partName,mediaPart:expected.mediaPart,relationshipId:expected.relationshipId,previousRelationshipId:expected.previousRelationshipId,previousMediaPart:expected.previousMediaPart});input.payload.fill(0);
 const after=archive(p.package.toBytes());expect([...after.keys()].sort()).toEqual([...before.keys(),expected.mediaPart].sort());custody(before,after,replacementContract.custody.allowedChangedParts);expect(after.get(expected.mediaPart)).toEqual(input.source);
 const a=selected(before,input.shapeId),b=selected(after,input.shapeId),attributeName=Object.keys(a.blip.attrs).find(k=>k.split(':').at(-1)==='embed'&&a.blip.bindings[k.split(':')[0]!]===R)!;
 const opening=a.source.slice(a.blip.start,a.blip.openEnd),match=new RegExp('('+attributeName+'\\s*=\\s*["\\\'])'+expected.previousRelationshipId+'(["\\\'])').exec(opening)!;expect(match).not.toBeNull();
 const changed=opening.replace(match[0],match[1]+expected.relationshipId+match[2]);expect(b.source).toBe(a.source.slice(0,a.blip.start)+changed+a.source.slice(a.blip.openEnd));
 const g=graph(after),oldGraph=graph(before);for(const edge of oldGraph.edges)expect(g.edges).toContainEqual(edge);expect(g.types.get(expected.mediaPart)).toBe(expected.contentType);
 expect(g.edges.find(e=>e.owner===s.partName&&e.id===expected.relationshipId)).toEqual({owner:s.partName,id:expected.relationshipId,type:R+'/image',target:expected.target,resolved:expected.mediaPart,external:false});
 const relPart='ppt/slides/_rels/slide1.xml.rels',rels=decode.decode(after.get(relPart)),node=xml(rels).children.find(n=>n.attrs.Id===expected.relationshipId)!;expect(rels.slice(0,node.start)+rels.slice(node.end)).toBe(decode.decode(before.get(relPart)));
 const next=s.inspectPictures();expect(next.map(r=>r.shapeId)).toEqual(previous.map(r=>r.shapeId));for(let n=0;n<previous.length;n++){const old=previous[n]!,current=next[n]!;if(old.shapeId!==input.shapeId)expect(current).toEqual(old);else{expect({...current,embedded:old.embedded}).toEqual(old);expect(current.embedded).toEqual({relationshipId:expected.relationshipId,target:expected.target,partName:expected.mediaPart,contentType:expected.contentType,byteLength:expected.byteLength});}}
 const dir=await mkdtemp(join(tmpdir(),'picture-replace-'));try{const path=join(dir,'saved.pptx');await p.save(path);const saved=await Bun.file(path).bytes();custody(after,archive(saved));expect((await Presentation.open(saved)).slides[0]!.inspectPictures()).toEqual(next);}finally{await rm(dir,{recursive:true,force:true});}
});
test('native replacement rollback after full graph and XML edit preserves version and retries cleanly',async()=>{
 const input=await replacementInput(replacementCases[1]),p=await Presentation.open(input.bytes),s=p.slides[0]!,before=archive(input.bytes),version=p.currentSlideVersion(s.partName),serialize=p.package.toBytes;let calls=0;
 p.package.toBytes=()=>{if(++calls===3)throw Error('replacement serialization failed');return serialize.call(p.package);};try{expect(()=>s.replacePicture(input.shapeId,input.payload,input.options)).toThrow('replacement serialization failed');}finally{p.package.toBytes=serialize;}
 const after=archive(p.package.toBytes());expect([...after.keys()].sort()).toEqual([...before.keys()].sort());custody(before,after);expect(p.currentSlideVersion(s.partName)).toBe(version);expect(s.replacePicture(input.shapeId,input.payload,input.options).relationshipId).toBe('rId4');
});
test('native replacement bounds and accessor refusal leave a reusable target',async()=>{
 const input=await replacementInput(replacementCases[1]),p=await Presentation.open(input.bytes),s=p.slides[0]!;let calls=0;
 for(const id of [NaN,0,-1,1.5,2147483648])expect(()=>s.replacePicture(id,input.payload,input.options)).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));
 const options={get contentType(){calls++;return 'image/jpeg';}};
 expect(()=>s.replacePicture(input.shapeId,input.payload,options as typeof input.options)).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));expect(calls).toBe(0);
 expect(()=>s.replacePicture(input.shapeId,input.payload,{...input.options,name:'no rename'} as typeof input.options)).toThrow(expect.objectContaining({code:'PPTX_PICTURE_UNSUPPORTED'}));expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});
 expect(s.replacePicture(input.shapeId,input.payload,input.options).shapeId).toBe(input.shapeId);
});
