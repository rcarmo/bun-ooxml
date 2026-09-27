import { OoxmlError } from '../errors.ts';
import { attribute, escapeAttribute, type XmlElement } from '../xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
/** Bounded editor policy; font existence and script coverage are not checked. */
export function fontNameValue(value:unknown,input:boolean):string {
 const code=input?'docx-format-argument':'docx-format-unsupported';
 if(typeof value!=='string'||!value.length||value.length>255||value.trim()!==value||/[\u0000-\u001f\u007f-\u009f]/.test(value))throw new OoxmlError(code,'Font name must be a nonempty, trimmed string of at most 255 UTF-16 code units without controls');
 try{escapeAttribute(value);}catch{throw new OoxmlError(code,'Font name contains invalid XML characters');}
 return value;
}
/** Read only a complete, matching pair of direct Latin font slots. */
export function fontNameFromNode(node:XmlElement):string {
 for(const name of Object.keys(node.attributes))if(node.attributeNamespaces[name]!==XMLNS&&!(node.attributeNamespaces[name]===W&&['ascii','hAnsi'].includes(name.split(':').at(-1)!)))throw new OoxmlError('docx-format-unsupported','Themed, mixed-script or hinted font metadata is unsupported');
 const ascii=fontNameValue(attribute(node,'ascii',W),false),hAnsi=fontNameValue(attribute(node,'hAnsi',W),false);
 if(ascii!==hAnsi)throw new OoxmlError('docx-format-unsupported','Direct Latin font slots must match');
 return ascii;
}
export function fontNameXml(value:string):string {
 const escaped=escapeAttribute(value);
 return `<w:rFonts xmlns:w="${W}" w:ascii="${escaped}" w:hAnsi="${escaped}"/>`;
}
