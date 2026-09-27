import {OoxmlError} from '../errors.ts';
import {applyEdits,attribute,escapeAttribute,type XmlElement} from '../xml/index.ts';
import {readRowHeader} from './row-header.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
const ORDER=['tblStyle','tblpPr','tblOverlap','bidiVisual','tblStyleRowBandSize','tblStyleColBandSize','tblW','jc','tblCellSpacing','tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription'];
const word=(n:XmlElement,name:string)=>n.namespaceURI===W&&n.localName===name;
function fail(message:string):never{throw new OoxmlError('docx-table-style-unsupported',message);}
function style(value:unknown):string{if(typeof value!=='string'||!value||value.trim()!==value||value.length>255||/[\u0000-\u001f\u007f-\u009f]/.test(value))fail('Style ID must be a nonblank bounded XML string without controls');escapeAttribute(value);return value;}
function gaps(xml:string,n:XmlElement){if(n.selfClosing)return;let pos=n.openEnd;for(const child of n.children){if(!/^[ \t\r\n]*$/.test(xml.slice(pos,child.start)))fail('Lexical table properties refuse');pos=child.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(pos,n.closeStart)))fail('Lexical table properties refuse');}
function inspect(xml:string,table:XmlElement){
 const rows=table.children.filter(n=>word(n,'tr')),grids=table.children.filter(n=>word(n,'tblGrid'));
 if(rows.length<1||rows.length>100||grids.length!==1||grids[0]!.children.length<1||grids[0]!.children.length>100)fail('Style editing requires a bounded explicit-grid table');
 for(let i=0;i<rows.length;i++)readRowHeader(xml,table,i);
 const pr=table.children.find(n=>word(n,'tblPr'));let node:XmlElement|undefined,value:string|undefined;
 if(pr){gaps(xml,pr);for(const key of Object.keys(pr.attributes))if(pr.attributeNamespaces[key]!==XMLNS)fail('Unsupported table-property attributes');const seen=new Set<string>();let last=-1;
  for(const child of pr.children){const rank=ORDER.indexOf(child.localName);if(child.namespaceURI!==W||rank<0||rank<last||seen.has(child.localName))fail('Unknown, duplicate or misordered table properties');last=rank;seen.add(child.localName);
   if(word(child,'tblStyle')){node=child;if(child.children.length)fail('Style reference must be a leaf');gaps(xml,child);for(const key of Object.keys(child.attributes))if(child.attributeNamespaces[key]!==XMLNS&&!(child.attributeNamespaces[key]===W&&key.split(':').at(-1)==='val'))fail('Malformed style attributes');value=style(attribute(child,'val',W));}
  }
 }
 return {pr,node,value};
}
export function readTableStyle(xml:string,table:XmlElement):string|undefined{return inspect(xml,table).value;}
export function editTableStyle(xml:string,table:XmlElement,input:string|null):string{
 const wanted=input===null?null:style(input),{pr,node,value}=inspect(xml,table);
 if(wanted===null&&!node||wanted===value)return xml;
 const fragment=wanted===null?'':`<w:tblStyle xmlns:w="${W}" w:val="${escapeAttribute(wanted)}"/>`;
 if(node)return applyEdits(xml,[{start:node.start,end:node.end,value:fragment}]);
 if(!pr)return applyEdits(xml,[{start:table.openEnd,end:table.openEnd,value:`<w:tblPr xmlns:w="${W}">${fragment}</w:tblPr>`}]);
 if(pr.selfClosing)return applyEdits(xml,[{start:pr.start,end:pr.end,value:xml.slice(pr.start,pr.end).replace(/\/>$/,'>')+fragment+`</${pr.name}>`}]);
 return applyEdits(xml,[{start:pr.openEnd,end:pr.openEnd,value:fragment}]);
}
