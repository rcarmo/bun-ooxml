import { OoxmlError } from '../errors.ts';
import { attribute, type XmlElement } from '../xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
export const APPEARANCE_NAMES={color:'color',underline:'u',highlight:'highlight',verticalAlign:'vertAlign'} as const;
export type AppearanceKey=keyof typeof APPEARANCE_NAMES;
export const APPEARANCE_KEYS=Object.keys(APPEARANCE_NAMES) as AppearanceKey[];
const UNDERLINES=['none','single','double','thick','dotted','dash','wave'] as const;
const HIGHLIGHTS=['black','blue','cyan','green','magenta','red','yellow','white','darkBlue','darkCyan','darkGreen','darkMagenta','darkRed','darkYellow','darkGray','lightGray','none'] as const;
const VERTICAL=['baseline','superscript','subscript'] as const;
export type UnderlineStyle=typeof UNDERLINES[number];
export type HighlightColor=typeof HIGHLIGHTS[number];
export type RunVerticalAlignment=typeof VERTICAL[number];
export interface DirectRunAppearance {color:string|null;underline:UnderlineStyle|null;highlight:HighlightColor|null;verticalAlign:RunVerticalAlignment|null}
export type RunAppearancePatch=Partial<DirectRunAppearance>;
export function appearanceValue(key:AppearanceKey,value:unknown,input:boolean):string {
 const refuse=():never=>{throw new OoxmlError(input?'docx-format-argument':'docx-format-unsupported',`Unsupported direct ${key} value`);};
 if(typeof value!=='string')return refuse();
 if(key==='color'){
  if(value==='auto')return value;
  const rgb=input&&value.startsWith('#')?value.slice(1):value;
  if(/^[0-9a-fA-F]{6}$/.test(rgb))return rgb;
  return refuse();
 }
 const values:readonly string[]=key==='underline'?UNDERLINES:key==='highlight'?HIGHLIGHTS:VERTICAL;
 if(!values.includes(value))return refuse();return value;
}
export function appearanceFromNode(key:AppearanceKey,node:XmlElement):string {
 for(const name of Object.keys(node.attributes))if(node.attributeNamespaces[name]!==XMLNS&&!(node.attributeNamespaces[name]===W&&name.split(':').at(-1)==='val'))throw new OoxmlError('docx-format-unsupported',`Decorated, themed or misqualified ${key} metadata is unsupported`);
 return appearanceValue(key,attribute(node,'val',W),false);
}
