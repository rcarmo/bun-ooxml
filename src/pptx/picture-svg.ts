import {posix} from 'node:path';
import {OoxmlError} from '../errors.ts';
import type {OpcPackage} from '../opc/package.ts';
import {addPart,addRelationship,nextPartName} from '../opc/graph.ts';
import {parseXml} from '../xml/index.ts';
import {pictureRequest,addPicture,type PictureGeometry,type PictureOptions,type PictureReceipt} from './picture-add.ts';
export type SvgPictureReceipt=PictureReceipt&{svgMediaPart:string;svgRelationshipId:string};
const S='http://www.w3.org/2000/svg',N='http://www.w3.org/2000/xmlns/',XML='http://www.w3.org/XML/1998/namespace',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
function fail(message:string):never {throw new OoxmlError('PPTX_PICTURE_UNSUPPORTED',message);}
export function svgPictureRequest(svg:Uint8Array,fallback:Uint8Array,geometry:PictureGeometry,options:PictureOptions){
 if(!(svg instanceof Uint8Array)||!svg.length||svg.length>1024*1024)fail('SVG must be nonempty UTF-8 bytes at most 1 MiB');
 const data=new Uint8Array(svg);let source:string,doc:ReturnType<typeof parseXml>;
 try{source=new TextDecoder('utf-8',{fatal:true}).decode(data);if(/<!DOCTYPE|<\?/i.test(source))fail('SVG declarations and processing instructions are unsupported');doc=parseXml(source);}catch{fail('Unsupported SVG XML');}
 if(doc.root.localName!=='svg'||doc.root.namespaceURI!==S)fail('Expected SVG namespace root');
 const tags=['svg','g','rect','circle','ellipse','line','polyline','polygon','path','title','desc'],attrs=['width','height','viewBox','preserveAspectRatio','x','y','x1','y1','x2','y2','cx','cy','r','rx','ry','points','d','fill','stroke','stroke-width','opacity','fill-opacity','stroke-opacity','fill-rule','stroke-linecap','stroke-linejoin','transform'];
 for(const node of doc.elements){
  if(node.namespaceURI!==S||!tags.includes(node.localName))fail('Unsupported SVG element');
  for(const [key,value]of Object.entries(node.attributes)){
   const ns=node.attributeNamespaces[key];if(ns===N)continue;
   if(!(ns===''&&attrs.includes(key)||ns===XML&&key.split(':').at(-1)==='space')||/url\s*\(/i.test(value))fail('Unsupported SVG attribute or paint reference');
   if(['fill','stroke'].includes(key)&&!/^(?:none|#[0-9a-fA-F]{3,8}|[a-zA-Z]+)$/.test(value))fail('SVG paints require plain colors or none');
  }
 }
 return {data,fallback:pictureRequest(fallback,geometry,options)};
}
/** Atomic paired authoring; no rasterisation or external reference resolution. */
export function addSvgPicture(pkg:OpcPackage,part:string,request:ReturnType<typeof svgPictureRequest>):SvgPictureReceipt{
 const svgMediaPart=nextPartName(pkg,'ppt/media/image%d.svg');
 return pkg.transaction(()=>{
  addPart(pkg,svgMediaPart,request.data,'image/svg+xml');
  const rel=addRelationship(pkg,part,R+'/image',posix.relative(posix.dirname(part),svgMediaPart));
  const receipt=addPicture(pkg,part,request.fallback,undefined,rel.id);
  return {...receipt,svgMediaPart,svgRelationshipId:rel.id};
 });
}
