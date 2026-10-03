import {OoxmlError} from '../errors.ts';
import type {OpcPackage} from '../opc/package.ts';
import {inspectPictures} from './pictures.ts';
import {parseXml,elements,attribute,applyEdits,type XmlElement} from '../xml/index.ts';
import {retainedFormattingXml} from './formatting.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',N='http://www.w3.org/2000/xmlns/';
export type PictureTransformPatch={rotation?:number;flipH?:boolean;flipV?:boolean};
function fail(message:string):never {throw new OoxmlError('PPTX_PICTURE_UNSUPPORTED',message);}
function patch(value:unknown):PictureTransformPatch{
 if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Picture orientation requires plain data');
 const copy:PictureTransformPatch={};
 for(const k of Reflect.ownKeys(value)){
  if(typeof k!=='string'||!['rotation','flipH','flipV'].includes(k))fail('Unknown orientation field');
  const d=Object.getOwnPropertyDescriptor(value,k)!;if(!('value'in d))fail('Orientation accessors are unsupported');
  if(k==='rotation'){if(!Number.isInteger(d.value)||d.value<0||d.value>21599999)fail('Rotation must be a bounded DrawingML angle');copy.rotation=d.value;}
  else{if(typeof d.value!=='boolean')fail('Picture flip requires a Boolean');copy[k as 'flipH'|'flipV']=d.value;}
 }
 return copy;
}
function attributes(node:XmlElement,allowed:string[],required:string[]=[]){
 const keys=Object.keys(node.attributes).filter(k=>node.attributeNamespaces[k]!==N);
 if(keys.some(k=>node.attributeNamespaces[k]!==''||!allowed.includes(k))||required.some(k=>!keys.includes(k)))fail('Unsupported picture transform attributes');
}
/** Edit only direct orientation attributes; coordinates/crop/assets remain literal. */
export function patchPictureTransformXml(pkg:OpcPackage,part:string,id:number,value:PictureTransformPatch):string{
 const changes=patch(value);if(!Number.isInteger(id)||id<1||id>2147483647)fail('Orientation requires a bounded picture ID');
 const info=inspectPictures(pkg,part).find(p=>p.shapeId===id);if(!info)throw new OoxmlError('PPTX_PICTURE_NOT_FOUND','No picture with the exact ID');
 if(!info.transform)fail('Picture has no direct transform');
 const source=pkg.text(part),doc=parseXml(source),picture=elements(doc,'pic',P).find(n=>n.children.find(c=>c.namespaceURI===P&&c.localName==='nvPicPr')?.children.some(c=>c.namespaceURI===P&&c.localName==='cNvPr'&&Number(attribute(c,'id'))===id))!,properties=picture.children.find(c=>c.namespaceURI===P&&c.localName==='spPr')!,transform=properties.children.find(c=>c.namespaceURI===A&&c.localName==='xfrm')!;
 attributes(transform,['rot','flipH','flipV']);
 if(info.transform.rotation<0||info.transform.rotation>21599999)fail('Retained rotation is outside bounded authoring profile');
 if(transform.children.length!==2||transform.children[0]!.namespaceURI!==A||transform.children[0]!.localName!=='off'||transform.children[1]!.namespaceURI!==A||transform.children[1]!.localName!=='ext')fail('Unsupported direct picture transform');
 let at=transform.openEnd;for(const c of transform.children){if(source.slice(at,c.start).trim()||c.children.length||(!c.selfClosing&&source.slice(c.openEnd,c.closeStart).trim()))fail('Picture transform lexical barrier');at=c.end;}if(source.slice(at,transform.closeStart).trim())fail('Picture transform lexical barrier');
 attributes(transform.children[0]!,['x','y'],['x','y']);attributes(transform.children[1]!,['cx','cy'],['cx','cy']);
 const edits:Record<string,string>={};
 if(changes.rotation!==undefined&&changes.rotation!==info.transform.rotation)edits.rot=String(changes.rotation);
 for(const key of ['flipH','flipV'] as const)if(changes[key]!==undefined&&changes[key]!==info.transform[key])edits[key]=changes[key]?'1':'0';
 if(!Object.keys(edits).length)return source;
 let opening:string;try{opening=retainedFormattingXml.openTag(source,transform,edits);}catch{fail('Unsupported orientation lexical tag');}
 return applyEdits(source,[{start:transform.start,end:transform.openEnd,value:opening}]);
}
