import {OoxmlError} from '../errors.ts';
import {parseXml,elements,applyEdits,escapeText,escapeAttribute,type XmlElement} from '../xml/index.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main';
const A='http://schemas.openxmlformats.org/drawingml/2006/main';
const MAX=2147483647;
export type TextBoxGeometry={x:number;y:number;width:number;height:number};
export type TextBoxOptions={name?:string;bold?:boolean;italic?:boolean};
export type TextBoxReceipt={shapeId:number;partName:string;paragraphCount:number};
const fail=(message:string):never=>{throw new OoxmlError('PPTX_TEXT_BOX_UNSUPPORTED',message);};
const is=(node:XmlElement,name:string,ns=P)=>node.namespaceURI===ns&&node.localName===name;
function plain(value:unknown,allowed:string[]):Record<string,unknown>{
 if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Options and geometry must be plain data');
 const copy:Record<string,unknown>={};
 for(const key of Reflect.ownKeys(value as object)){
  if(typeof key!=='string'||!allowed.includes(key))fail('Unknown text-box property');
  const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Accessor properties are unsupported');copy[key as string]=d.value;
 }
 return copy;
}
export function textBoxRequest(text:string,geometry:TextBoxGeometry,options:TextBoxOptions){
 if(typeof text!=='string')fail('Text must be a string');escapeText(text);
 const g=plain(geometry,['x','y','width','height']);
 for(const key of ['x','y','width','height']){const v=g[key];if(typeof v!=='number'||!Number.isInteger(v)||v<(key==='x'||key==='y'?0:1)||v>MAX)fail('Geometry requires bounded nonnegative integer positions and positive extents');}
 const o=plain(options,['name','bold','italic']);
 if(o.name!==undefined){if(typeof o.name!=='string'||!o.name.trim())fail('Shape name must be nonempty text');escapeAttribute(o.name as string);}
 for(const key of ['bold','italic'])if(o[key]!==undefined&&typeof o[key]!=='boolean')fail('Bold and italic must be booleans');
 return {text:text.replace(/\r\n?/g,'\n'),geometry:g as TextBoxGeometry,options:o as TextBoxOptions};
}
function unique(parent:XmlElement,name:string):XmlElement{
 const nodes=parent.children.filter(n=>is(n,name));if(nodes.length!==1)fail(`Expected exactly one ${name}`);return nodes[0]!;
}
function whitespace(xml:string,node:XmlElement){
 if(node.selfClosing)return;let cursor=node.openEnd;
 for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,child.start)))fail('Shape-tree lexical barriers are unsupported');cursor=child.end;}
 if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,node.closeStart)))fail('Shape-tree lexical barriers are unsupported');
}
function validateIdentities(tree:XmlElement){
 for(const child of tree.children){
  const nv=({sp:'nvSpPr',grpSp:'nvGrpSpPr',graphicFrame:'nvGraphicFramePr',cxnSp:'nvCxnSpPr',pic:'nvPicPr'} as Record<string,string>)[child.localName];
  if(child.namespaceURI===P&&nv){unique(unique(child,nv),'cNvPr');if(is(child,'grpSp'))validateIdentities(child);}
 }
}
/** Append only; no existing shape, relationship, placeholder or master is rewritten. */
export function appendTextBox(xml:string,request:ReturnType<typeof textBoxRequest>):{xml:string;shapeId:number;paragraphCount:number}{
 const doc=parseXml(xml);if(!is(doc.root,'sld'))fail('Invalid slide root');const tree=unique(unique(doc.root,'cSld'),'spTree');
 whitespace(xml,tree);
 if(tree.children.length<2||!is(tree.children[0]!,'nvGrpSpPr')||!is(tree.children[1]!,'grpSpPr'))fail('Shape tree requires leading group properties');
 const nv=unique(tree,'nvGrpSpPr'),props=unique(tree,'grpSpPr');unique(nv,'cNvPr');
 if(tree.children.some((n,i)=>n.namespaceURI!==P||!['nvGrpSpPr','grpSpPr','sp','grpSp','graphicFrame','cxnSp','pic','extLst'].includes(n.localName)||is(n,'extLst')&&i!==tree.children.length-1))fail('Unsupported or misordered shape tree child');
 // Non-identity root transforms would change the meaning of slide-space EMUs.
 const transforms=props.children.filter(n=>is(n,'xfrm',A));if(props.children.length!==transforms.length||transforms.length>1)fail('Unsupported root group properties');
 if(transforms.length){const x=transforms[0]!;if(Object.keys(x.attributes).length)fail('Root transform attributes are unsupported');const names=['off','ext','chOff','chExt'];if(x.children.length!==4||x.children.some((n,i)=>!is(n,names[i]!,A)))fail('Invalid root transform');for(const n of x.children){const keys=n.localName.endsWith('Off')||n.localName==='off'?['x','y']:['cx','cy'];if(n.children.length||Object.keys(n.attributes).length!==2||keys.some(k=>!/^0+$/.test(n.attributes[k]??'')))fail('Nonzero root transform is unsupported');}}
 validateIdentities(tree);
 const seen=new Set<number>();let maximum=0;
 for(const node of elements(doc,'cNvPr',P)){const raw=node.attributes.id;if(!raw||!/^\d+$/.test(raw))fail('Malformed shape ID');const n=Number(raw);if(!Number.isSafeInteger(n)||n<1||n>MAX||seen.has(n))fail('Invalid or duplicate shape ID');seen.add(n);maximum=Math.max(maximum,n);}
 if(maximum>=MAX)throw new OoxmlError('PPTX_ID_EXHAUSTED','No next shape ID available');
 const id=maximum+1,g=request.geometry,o=request.options,lines=request.text.split('\n');
 const flags=(['bold','italic'] as const).filter(k=>o[k]!==undefined).map(k=>` ${k==='bold'?'b':'i'}="${o[k]?'1':'0'}"`).join('');
 const shape=`<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:nvSpPr><p:cNvPr id="${id}" name="${escapeAttribute(o.name??`TextBox ${id}`)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${g.x}" y="${g.y}"/><a:ext cx="${g.width}" cy="${g.height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr><a:lstStyle/>${lines.map(t=>`<a:p><a:r>${flags?`<a:rPr${flags}/>`:''}<a:t xml:space="preserve">${escapeText(t)}</a:t></a:r></a:p>`).join('')}</p:txBody></p:sp>`;
 const at=tree.children.find(n=>is(n,'extLst'))?.start??tree.closeStart;
 return {xml:applyEdits(xml,[{start:at,end:at,value:shape}]),shapeId:id,paragraphCount:lines.length};
}
