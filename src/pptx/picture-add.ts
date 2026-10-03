import {posix} from 'node:path';
import {OoxmlError} from '../errors.ts';
import type {OpcPackage} from '../opc/package.ts';
import {addPart,addRelationship,nextPartName} from '../opc/graph.ts';
import {applyEdits,escapeAttribute} from '../xml/index.ts';
import {shapeAppendSite} from './text-box.ts';
import type {PictureCrop} from './pictures.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export type PictureGeometry={x:number;y:number;width:number;height:number};
export type PictureOptions={contentType:'image/png'|'image/jpeg';name?:string;description?:string};
export type PictureReceipt={shapeId:number;partName:string;mediaPart:string;relationshipId:string};
function fail(message:string):never {throw new OoxmlError('PPTX_PICTURE_UNSUPPORTED',message);}
function plain(value:unknown,keys:string[]):Record<string,unknown>{
 if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Picture request requires plain data');
 const result:Record<string,unknown>={};
 for(const key of Reflect.ownKeys(value)){if(typeof key!=='string'||!keys.includes(key))fail('Unknown picture request property');const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Picture accessors are unsupported');result[key]=d.value;}
 return result;
}
export function pictureRequest(bytes:Uint8Array,geometry:PictureGeometry,options:PictureOptions){
 if(!(bytes instanceof Uint8Array)||bytes.length===0||bytes.length>64*1024*1024)fail('Picture bytes must be nonempty and at most 64 MiB');
 const data=new Uint8Array(bytes),g=plain(geometry,['x','y','width','height']),o=plain(options,['contentType','name','description']);
 const png=[137,80,78,71,13,10,26,10].every((b,i)=>data[i]===b),jpeg=data.length>=4&&data[0]===255&&data[1]===216&&data.at(-2)===255&&data.at(-1)===217;
 if(!(o.contentType==='image/png'&&png||o.contentType==='image/jpeg'&&jpeg))fail('Unsupported picture MIME or signature mismatch');
 for(const key of ['x','y','width','height']){const v=g[key];if(typeof v!=='number'||!Number.isInteger(v)||v<(key==='x'||key==='y'?-2147483648:1)||v>2147483647)fail('Picture rectangle requires bounded integer coordinates and positive extents');}
 for(const key of ['name','description'])if(o[key]!==undefined){if(typeof o[key]!=='string'||key==='name'&&!(o[key] as string).trim())fail('Picture name/description must be text');try{escapeAttribute(o[key] as string);}catch{fail('Invalid picture XML string');}}
 return {data,geometry:g as PictureGeometry,options:o as PictureOptions};
}
/** All graph and lexical edits share one rollback boundary; payloads are never decoded. */
export function addPicture(pkg:OpcPackage,part:string,request:ReturnType<typeof pictureRequest>,crop?:PictureCrop,svgRelationshipId?:string):PictureReceipt{
 const xml=pkg.text(part);let site:ReturnType<typeof shapeAppendSite>;
 try{site=shapeAppendSite(xml);}catch(error){if(error instanceof OoxmlError&&error.code==='PPTX_TEXT_BOX_UNSUPPORTED')fail(error.message);throw error;}
 const {shapeId,at}=site,g=request.geometry,o=request.options;
 const mediaPart=nextPartName(pkg,`ppt/media/image%d.${o.contentType==='image/png'?'png':'jpeg'}`);
 return pkg.transaction(()=>{
  addPart(pkg,mediaPart,request.data,o.contentType);
  const rel=addRelationship(pkg,part,R+'/image',posix.relative(posix.dirname(part),mediaPart));
  const picture=`<p:pic xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:nvPicPr><p:cNvPr id="${shapeId}" name="${escapeAttribute(o.name??`Picture ${shapeId}`)}"${o.description!==undefined?` descr="${escapeAttribute(o.description)}"`:''}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${escapeAttribute(rel.id)}"${svgRelationshipId?`><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${escapeAttribute(svgRelationshipId)}"/></a:ext></a:extLst></a:blip>`:'/>'}${crop&&Object.values(crop).some(v=>v!==0)?`<a:srcRect l="${crop.left}" t="${crop.top}" r="${crop.right}" b="${crop.bottom}"/>`:''}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${g.x}" y="${g.y}"/><a:ext cx="${g.width}" cy="${g.height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  pkg.set(part,applyEdits(xml,[{start:at,end:at,value:picture}]));pkg.toBytes();
  return {shapeId,partName:part,mediaPart,relationshipId:rel.id};
 });
}
