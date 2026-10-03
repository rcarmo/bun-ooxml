import {OoxmlError} from '../errors.ts';
import {parseXml,elements,attribute,escapeAttribute,applyEdits,type XmlElement} from '../xml/index.ts';
import {shapeAppendSite} from './text-box.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',N='http://www.w3.org/2000/xmlns/',MAX=2147483647;
export type ConnectorEndpoint={shapeId:number;site:number};
export type ConnectorOptions={name?:string;color?:string;width?:number};
export type ConnectorPoint=ConnectorEndpoint&{x:number;y:number};
export type ConnectorGeometry={x:number;y:number;width:number;height:number;flipH:boolean;flipV:boolean};
export type ConnectorReceipt={shapeId:number;partName:string;start:ConnectorPoint;end:ConnectorPoint;geometry:ConnectorGeometry};
function fail(message:string):never{throw new OoxmlError('PPTX_CONNECTOR_UNSUPPORTED',message);}
function plain(value:unknown,allowed:string[]):Record<string,unknown>{if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Connector inputs require plain data');const copy:Record<string,unknown>={};for(const key of Reflect.ownKeys(value)){if(typeof key!=='string'||!allowed.includes(key))fail('Unknown connector field');const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Connector accessors are unsupported');copy[key]=d.value;}return copy;}
function bounded(value:unknown,min:number,max=MAX):number{if(typeof value!=='number'||!Number.isInteger(value)||value<min||value>max)fail('Connector value must be a bounded integer');return value;}
function endpoint(value:unknown):ConnectorEndpoint{const e=plain(value,['shapeId','site']);return {shapeId:bounded(e.shapeId,1),site:bounded(e.site,0,3)};}
function one(parent:XmlElement,name:string,ns:string):XmlElement{const nodes=parent.children.filter(n=>n.namespaceURI===ns&&n.localName===name);if(nodes.length!==1)fail('Expected one '+name);return nodes[0]!;}
function attrs(node:XmlElement,allowed:string[],required:string[]=[]){const keys=Object.keys(node.attributes).filter(k=>node.attributeNamespaces[k]!==N);if(keys.some(k=>node.attributeNamespaces[k]!==''||!allowed.includes(k))||required.some(k=>!keys.includes(k)))fail('Unsupported endpoint attributes');}
function integer(node:XmlElement,key:string,min:number){const raw=attribute(node,key);if(raw===undefined||!/^[-+]?\d+$/.test(raw))fail('Invalid endpoint integer');return bounded(Number(raw),min);}
function site(source:string,tree:XmlElement,e:ConnectorEndpoint):ConnectorPoint{
 const nodes=tree.children.filter(n=>n.namespaceURI===P&&n.localName==='sp'&&Number(attribute(one(one(n,'nvSpPr',P),'cNvPr',P),'id'))===e.shapeId);if(nodes.length!==1)fail('Missing direct rectangle endpoint');const shape=nodes[0]!;
 if(elements(shape,'ph',P).length)fail('Placeholder endpoints are unsupported');for(const lock of elements(shape,'spLocks',A))if(![undefined,'0','false'].includes(attribute(lock,'noConnect')))fail('Endpoint connection is locked');
 const props=one(shape,'spPr',P),transform=one(props,'xfrm',A),preset=one(props,'prstGeom',A);attrs(preset,['prst'],['prst']);if(attribute(preset,'prst')!=='rect'||preset.children.length!==1)fail('Endpoint requires a rectangle preset');const adjustments=one(preset,'avLst',A);attrs(adjustments,[]);if(adjustments.children.length||adjustments.directText.trim())fail('Endpoint adjustments are unsupported');
 attrs(transform,['rot','flipH','flipV']);if(attribute(transform,'rot')!==undefined&&integer(transform,'rot',0)!==0)fail('Rotated endpoints are unsupported');for(const key of ['flipH','flipV'])if(![undefined,'0','false'].includes(attribute(transform,key)))fail('Flipped endpoints are unsupported');
 const off=one(transform,'off',A),ext=one(transform,'ext',A);if(transform.children.length!==2||transform.children[0]!==off||transform.children[1]!==ext)fail('Ambiguous endpoint transform');attrs(off,['x','y'],['x','y']);attrs(ext,['cx','cy'],['cx','cy']);let cursor=transform.openEnd;for(const n of transform.children){if(source.slice(cursor,n.start).trim()||n.children.length||n.directText.trim())fail('Endpoint transform lexical barrier');cursor=n.end;}if(source.slice(cursor,transform.closeStart).trim())fail('Endpoint transform lexical barrier');
 const x=integer(off,'x',-2147483648),y=integer(off,'y',-2147483648),w=integer(ext,'cx',1),h=integer(ext,'cy',1),points=[[x+Math.floor(w/2),y],[x,y+Math.floor(h/2)],[x+Math.floor(w/2),y+h],[x+w,y+Math.floor(h/2)]];
 const chosen=points[e.site]!;return {...e,x:bounded(chosen[0],-2147483648),y:bounded(chosen[1],-2147483648)};
}
/** Append a direct straight connector; no routing guesses or endpoint rewriting. */
export function addConnectorXml(source:string,start:ConnectorEndpoint,end:ConnectorEndpoint,options:ConnectorOptions={}){
 const a=endpoint(start),b=endpoint(end),o=plain(options,['name','color','width']);if(a.shapeId===b.shapeId)fail('Connector endpoints require distinct shapes');
 if(o.name!==undefined){if(typeof o.name!=='string'||!o.name.trim())fail('Connector name must be nonempty');try{escapeAttribute(o.name);}catch{fail('Invalid connector name');}}
 const color=o.color??'000000';if(typeof color!=='string'||!/^([0-9A-F]{6})$/.test(color))fail('Connector colour requires uppercase RGB');const width=bounded(o.width??12700,1,20116800);
 let insertion:ReturnType<typeof shapeAppendSite>;try{insertion=shapeAppendSite(source);}catch(error){if(error instanceof OoxmlError&&error.code==='PPTX_TEXT_BOX_UNSUPPORTED')fail(error.message);throw error;}
 const doc=parseXml(source),tree=elements(doc,'spTree',P)[0]!,first=site(source,tree,a),last=site(source,tree,b),geometry={x:Math.min(first.x,last.x),y:Math.min(first.y,last.y),width:bounded(Math.abs(last.x-first.x),0),height:bounded(Math.abs(last.y-first.y),0),flipH:last.x<first.x,flipV:last.y<first.y};if(!geometry.width&&!geometry.height)fail('Coincident connector endpoints');
 const id=insertion.shapeId,shape=`<p:cxnSp xmlns:p="${P}" xmlns:a="${A}"><p:nvCxnSpPr><p:cNvPr id="${id}" name="${escapeAttribute(o.name as string??`Connector ${id}`)}"/><p:cNvCxnSpPr><a:stCxn id="${a.shapeId}" idx="${a.site}"/><a:endCxn id="${b.shapeId}" idx="${b.site}"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm flipH="${geometry.flipH?'1':'0'}" flipV="${geometry.flipV?'1':'0'}"><a:off x="${geometry.x}" y="${geometry.y}"/><a:ext cx="${geometry.width}" cy="${geometry.height}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="${width}"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`;
 return {xml:applyEdits(source,[{start:insertion.at,end:insertion.at,value:shape}]),shapeId:id,start:first,end:last,geometry};
}
