import {posix} from 'node:path';
import {OoxmlError} from '../errors.ts';
import type {OpcPackage} from '../opc/package.ts';
import {addPart,addRelationship,nextPartName} from '../opc/graph.ts';
import {attribute,parseXml,elements,applyEdits} from '../xml/index.ts';
import {inspectPictures} from './pictures.ts';
import {pictureRequest,type PictureReceipt} from './picture-add.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export type PictureReplacementOptions={contentType:'image/png'|'image/jpeg'};
export type PictureReplacementReceipt=PictureReceipt&{previousRelationshipId:string;previousMediaPart:string};
function fail(message:string):never {throw new OoxmlError('PPTX_PICTURE_UNSUPPORTED',message);}
export function replacementRequest(shapeId:number,bytes:Uint8Array,options:PictureReplacementOptions){
 if(!Number.isInteger(shapeId)||shapeId<1||shapeId>2147483647)fail('Replacement requires a bounded exact picture ID');
 if(!options||typeof options!=='object'||Reflect.ownKeys(options).some(k=>k!=='contentType'))fail('Replacement accepts only contentType');
 const request=pictureRequest(bytes,{x:0,y:0,width:1,height:1},options);return {shapeId,data:request.data,contentType:request.options.contentType};
}
/** Retarget exactly one embedded attribute; never overwrite/collect shared dependencies. */
export function replacePicture(pkg:OpcPackage,part:string,request:ReturnType<typeof replacementRequest>):PictureReplacementReceipt{
 const info=inspectPictures(pkg,part).find(p=>p.shapeId===request.shapeId);
 if(!info)throw new OoxmlError('PPTX_PICTURE_NOT_FOUND','No picture with the exact shape ID');
 if(!info.embedded||info.linked)fail('Replacement requires an embedded-only picture');
 const source=pkg.text(part),doc=parseXml(source),picture=elements(doc,'pic',P).find(n=>n.children.find(c=>c.localName==='nvPicPr'&&c.namespaceURI===P)?.children.some(c=>c.localName==='cNvPr'&&c.namespaceURI===P&&Number(attribute(c,'id'))===request.shapeId))!;
 const fill=picture.children.find(n=>n.localName==='blipFill'&&n.namespaceURI===P)!,blip=fill.children.find(n=>n.localName==='blip'&&n.namespaceURI===A)!;
 // A paired extension asset (e.g. SVG fallback) needs dependency-aware replacement.
 function paired(node:typeof blip):boolean {return node.children.some(n=>attribute(n,'embed',R)!==undefined||attribute(n,'link',R)!==undefined||paired(n));}
 if(paired(blip))fail('Paired image extension replacement is unsupported');
 const key=Object.keys(blip.attributes).find(k=>k.split(':').at(-1)==='embed'&&blip.attributeNamespaces[k]===R)!;
 const opening=source.slice(blip.start,blip.openEnd),tokens=/\s+([^\s=/>]+)\s*=\s*(["'])([\s\S]*?)\2/g;
 const matches=[...opening.matchAll(tokens)].filter(m=>m[1]===key);
 if(matches.length!==1)fail('Embedded attribute has no unique lexical value');
 const match=matches[0]!,offset=match.index!+match[0].indexOf(match[2]!)+1,start=blip.start+offset,end=start+match[3]!.length;
 const mediaPart=nextPartName(pkg,`ppt/media/image%d.${request.contentType==='image/png'?'png':'jpeg'}`),previous=info.embedded;
 return pkg.transaction(()=>{
  addPart(pkg,mediaPart,request.data,request.contentType);const rel=addRelationship(pkg,part,R+'/image',posix.relative(posix.dirname(part),mediaPart));
  pkg.set(part,applyEdits(source,[{start,end,value:rel.id}]));pkg.toBytes();
  return {shapeId:request.shapeId,partName:part,mediaPart,relationshipId:rel.id,previousRelationshipId:previous.relationshipId,previousMediaPart:previous.partName};
 });
}
