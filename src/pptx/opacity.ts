import {OoxmlError} from '../errors.ts';
import type {OpcPackage} from '../opc/package.ts';
import {parseXml,elements,attribute,applyEdits,type XmlElement} from '../xml/index.ts';
import {getLinearGradientXml} from './gradients.ts';
import {inspectPictures} from './pictures.ts';
import {retainedFormattingXml} from './formatting.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',N='http://www.w3.org/2000/xmlns/';
function fail(message:string):never{throw new OoxmlError('PPTX_OPACITY_UNSUPPORTED',message);}
function scalar(value:unknown):number{if(typeof value!=='number'||!Number.isInteger(value)||value<0||value>100000)fail('Alpha profile requires an integer in [0,100000]');return value;}
function validId(id:number){if(!Number.isInteger(id)||id<1||id>2147483647)fail('Alpha profile requires a bounded shape ID');}
function gaps(source:string,n:XmlElement){if(n.selfClosing)return;let at=n.openEnd;for(const c of n.children){if(source.slice(at,c.start).trim())fail('Alpha profile lexical barrier');at=c.end;}if(source.slice(at,n.closeStart).trim())fail('Alpha profile lexical barrier');}
function effect(source:string,n:XmlElement,key:string):number{const keys=Object.keys(n.attributes).filter(k=>n.attributeNamespaces[k]!==N);if(n.children.length||keys.length!==1||keys[0]!==key||n.attributeNamespaces[key]!==''||!/^\d+$/.test(attribute(n,key)??''))fail('Unsupported alpha effect attributes');gaps(source,n);return scalar(Number(attribute(n,key)));}
function shape(source:string,id:number){
 validId(id);try{if(getLinearGradientXml(source,id)!==null)fail('Shape opacity requires a solid fill');}catch(error){if(error instanceof OoxmlError&&error.code==='PPTX_GRADIENT_UNSUPPORTED')fail(error.message);throw error;}
 const doc=parseXml(source),node=elements(doc,'sp',P).find(n=>n.children.find(c=>c.localName==='nvSpPr'&&c.namespaceURI===P)?.children.some(c=>c.localName==='cNvPr'&&c.namespaceURI===P&&Number(attribute(c,'id'))===id))!,pr=node.children.find(n=>n.localName==='spPr'&&n.namespaceURI===P)!,fills=pr.children.filter(n=>n.localName==='solidFill'&&n.namespaceURI===A);if(fills.length!==1)fail('Shape opacity requires a direct solid fill');const color=fills[0]!.children[0]!,rows=color.children.filter(n=>n.localName==='alpha'&&n.namespaceURI===A);if(rows.length>1)fail('Duplicate shape alpha');const alpha=rows[0],value=alpha?effect(source,alpha,'val'):100000;return {owner:color,alpha,value};
}
function picture(pkg:OpcPackage,part:string,id:number){
 validId(id);if(!inspectPictures(pkg,part).some(n=>n.shapeId===id))fail('Missing picture alpha target');const source=pkg.text(part),doc=parseXml(source),node=elements(doc,'pic',P).find(n=>n.children.find(c=>c.localName==='nvPicPr'&&c.namespaceURI===P)?.children.some(c=>c.localName==='cNvPr'&&c.namespaceURI===P&&Number(attribute(c,'id'))===id))!,fill=node.children.find(n=>n.localName==='blipFill'&&n.namespaceURI===P)!,blip=fill.children.find(n=>n.localName==='blip'&&n.namespaceURI===A)!;
 const order=['alphaModFix','biLevel','blur','clrChange','clrRepl','duotone','fillOverlay','grayscl','hsl','lum','tint','extLst'];let last=-1;gaps(source,blip);for(const n of blip.children){const i=order.indexOf(n.localName);if(n.namespaceURI!==A||i<0||i<=last)fail('Unsupported/competing picture effect grammar');last=i;}const alpha=blip.children.find(n=>n.localName==='alphaModFix'),value=alpha?100000-effect(source,alpha,'amt'):0;return {source,owner:blip,alpha,value};
}
function splice(source:string,owner:XmlElement,alpha:XmlElement|undefined,key:string,value:number,tag:string,at=owner.closeStart){
 if(alpha){let opening:string;try{opening=retainedFormattingXml.openTag(source,alpha,{[key]:String(value)});}catch{fail('Unsupported alpha lexical tag');}return applyEdits(source,[{start:alpha.start,end:alpha.openEnd,value:opening}]);}
 const child=`<a:${tag} xmlns:a="${A}" ${key}="${value}"/>`;
 if(owner.selfClosing)return applyEdits(source,[{start:owner.start,end:owner.end,value:source.slice(owner.start,owner.openEnd).replace(/\/\s*>$/,()=>`>${child}</${owner.name}>`)}]);
 return applyEdits(source,[{start:at,end:at,value:child}]);
}
export function getShapeOpacityXml(source:string,id:number):number{return shape(source,id).value;}
export function setShapeOpacityXml(source:string,id:number,input:number):string{const value=scalar(input),current=shape(source,id);return value===current.value?source:splice(source,current.owner,current.alpha,'val',value,'alpha');}
export function getPictureTransparency(pkg:OpcPackage,part:string,id:number):number{return picture(pkg,part,id).value;}
export function setPictureTransparencyXml(pkg:OpcPackage,part:string,id:number,input:number):string{const value=scalar(input),current=picture(pkg,part,id);return value===current.value?current.source:splice(current.source,current.owner,current.alpha,'amt',100000-value,'alphaModFix',current.owner.children[0]?.start??current.owner.closeStart);}
