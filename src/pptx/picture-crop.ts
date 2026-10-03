import {OoxmlError} from '../errors.ts';
import type {OpcPackage} from '../opc/package.ts';
import {inspectPictures,type PictureCrop} from './pictures.ts';
import {parseXml,elements,attribute,applyEdits,type XmlElement} from '../xml/index.ts';
import {retainedFormattingXml} from './formatting.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',N='http://www.w3.org/2000/xmlns/';
function fail(message:string):never {throw new OoxmlError('PPTX_PICTURE_UNSUPPORTED',message);}
function bounded(value:unknown):PictureCrop{
 if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Crop requires plain data');
 const keys=['left','top','right','bottom'],copy:Record<string,number>={};
 for(const k of Reflect.ownKeys(value)){if(typeof k!=='string'||!keys.includes(k))fail('Unknown crop field');const d=Object.getOwnPropertyDescriptor(value,k)!;if(!('value'in d))fail('Crop accessors are unsupported');if(!Number.isInteger(d.value)||d.value<0||d.value>99999)fail('Crop sides require bounded nonnegative integers');copy[k]=d.value;}
 if(keys.some(k=>copy[k]===undefined)||copy.left!+copy.right!>=100000||copy.top!+copy.bottom!>=100000)fail('Crop requires four sides and a nonempty visible region');
 return copy as PictureCrop;
}
function selected(pkg:OpcPackage,part:string,id:number){
 if(!Number.isInteger(id)||id<1||id>2147483647)fail('Crop requires a bounded picture ID');
 const info=inspectPictures(pkg,part).find(p=>p.shapeId===id);if(!info)throw new OoxmlError('PPTX_PICTURE_NOT_FOUND','No picture with the exact ID');
 const source=pkg.text(part),doc=parseXml(source),picture=elements(doc,'pic',P).find(n=>n.children.find(c=>c.namespaceURI===P&&c.localName==='nvPicPr')?.children.some(c=>c.namespaceURI===P&&c.localName==='cNvPr'&&Number(attribute(c,'id'))===id))!,fill=picture.children.find(c=>c.namespaceURI===P&&c.localName==='blipFill')!;
 const names=['blip','srcRect','tile','stretch','extLst'];let last=-1;const seen=new Set<string>();
 for(const n of fill.children){const at=names.indexOf(n.localName);if(n.namespaceURI!==A||at<0||at<=last||seen.has(n.localName))fail('Unsupported picture fill order');seen.add(n.localName);last=at;}
 if(seen.has('tile')&&seen.has('stretch'))fail('Ambiguous fill policy');
 const rect=fill.children.find(c=>c.namespaceURI===A&&c.localName==='srcRect');
 if(rect){if(rect.children.length||(!rect.selfClosing&&source.slice(rect.openEnd,rect.closeStart).trim()))fail('Mixed crop node');for(const k of Object.keys(rect.attributes))if(rect.attributeNamespaces[k]!==N&&(rect.attributeNamespaces[k]!==''||!['l','t','r','b'].includes(k)))fail('Unknown or foreign crop attribute');}
 return {source,fill,rect,crop:bounded(info.crop)};
}
export function getPictureCrop(pkg:OpcPackage,part:string,id:number):PictureCrop{return {...selected(pkg,part,id).crop};}
export function setPictureCropXml(pkg:OpcPackage,part:string,id:number,value:PictureCrop):string{
 const crop=bounded(value),current=selected(pkg,part,id);
 if((['left','top','right','bottom'] as const).every(k=>crop[k]===current.crop[k]))return current.source;
 const {source,fill,rect}=current,attrs={l:String(crop.left),t:String(crop.top),r:String(crop.right),b:String(crop.bottom)};
 if(rect){let opening:string;try{opening=retainedFormattingXml.openTag(source,rect,attrs);}catch{fail('Unsupported crop lexical tag');}return applyEdits(source,[{start:rect.start,end:rect.openEnd,value:opening}]);}
 const at=fill.children.find(n=>n.localName!=='blip')?.start??fill.closeStart;
 return applyEdits(source,[{start:at,end:at,value:`<a:srcRect xmlns:a="${A}" l="${attrs.l}" t="${attrs.t}" r="${attrs.r}" b="${attrs.b}"/>`}]);
}
