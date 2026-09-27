import {OoxmlError} from '../errors.ts';
import {applyEdits,attribute,type XmlElement} from '../xml/index.ts';
import {readCellProperties} from './cell-properties.ts';
import {formatRunProperties} from './run-formatting.ts';
import {directParagraphStyle} from './paragraph-style.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS='http://www.w3.org/2000/xmlns/';
const META:Record<string,readonly string[]>={cnfStyle:['val','firstRow','lastRow','firstColumn','lastColumn','oddVBand','evenVBand','oddHBand','evenHBand','firstRowFirstColumn','firstRowLastColumn','lastRowFirstColumn','lastRowLastColumn'],divId:['val'],gridBefore:['val'],gridAfter:['val'],wBefore:['w','type'],wAfter:['w','type'],cantSplit:['val'],trHeight:['val','hRule'],tblHeader:['val'],tblCellSpacing:['w','type'],jc:['val'],hidden:['val']};
function fail(message:string):never {throw new OoxmlError('docx-row-header-unsupported',message);}
const word=(node:XmlElement,name:string)=>node.namespaceURI===W&&node.localName===name;
function noMixed(xml:string,node:XmlElement):void{
 if(node.selfClosing)return;let pos=node.openEnd;for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(pos,child.start)))fail('Mixed or lexical row content');pos=child.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(pos,node.closeStart)))fail('Mixed or lexical row content');
}
function attrs(node:XmlElement,allowed:readonly string[]):void{for(const name of Object.keys(node.attributes))if(node.attributeNamespaces[name]!==XMLNS&&!(node.attributeNamespaces[name]===W&&allowed.includes(name.split(':').at(-1)!)))fail('Unsupported row-property attributes');}
function noRevision(xml:string,node:XmlElement):void{
 noMixed(xml,node);if(node.namespaceURI!==W||node.localName.endsWith('Change')||['ins','del','cellIns','cellDel','cellMerge'].includes(node.localName))fail('Unsupported or revised table metadata');for(const child of node.children)noRevision(xml,child);
}
function inspect(xml:string,table:XmlElement,index:number){
 noMixed(xml,table);let previous=-1;const seen=new Set<string>();
 for(const child of table.children){const rank=['tblPr','tblGrid','tr'].indexOf(child.localName);if(child.namespaceURI!==W||rank<previous||rank<0||(rank<2&&seen.has(child.localName)))fail('Unsupported table property order');previous=rank;seen.add(child.localName);if(rank<2)noRevision(xml,child);}
 if(!seen.has('tblGrid'))fail('Row headers require an explicit table grid');
 const row=table.children.filter(n=>word(n,'tr'))[index];if(!row)fail('Missing table row');noMixed(xml,row);
 let pr:XmlElement|undefined,haveCell=false;
 for(const child of row.children){if(word(child,'trPr')){if(pr||haveCell)fail('Duplicate or misplaced row properties');pr=child;}else if(word(child,'tc')){
  haveCell=true;noMixed(xml,child);readCellProperties(xml,child);
  for(const paragraph of child.children){if(word(paragraph,'tcPr'))continue;if(!word(paragraph,'p'))fail('Only plain cell paragraphs are supported');formatRunProperties(xml,paragraph,{});directParagraphStyle(xml,paragraph);}
 }else fail('Unsupported row topology');}
 if(!haveCell)fail('Missing row cells');
 let marker:XmlElement|undefined,value:boolean|null=null;
 if(pr){attrs(pr,[]);noMixed(xml,pr);const names=new Set<string>();for(const child of pr.children){
  const allowed=Object.hasOwn(META,child.localName)?META[child.localName]:undefined;
  if(child.namespaceURI!==W||!allowed||names.has(child.localName)||child.children.length)fail('Unsupported, duplicate or revised row properties');names.add(child.localName);attrs(child,allowed);noMixed(xml,child);
  if(word(child,'tblHeader')){marker=child;const raw=attribute(child,'val',W);if(raw===undefined||['1','true','on'].includes(raw))value=true;else if(['0','false','off'].includes(raw))value=false;else fail('Invalid row-header Boolean');}
 }}
 return {row,pr,marker,value};
}
export function readRowHeader(xml:string,table:XmlElement,row:number):boolean{return inspect(xml,table,row).value??false;}
export function editRowHeader(xml:string,table:XmlElement,row:number,value:boolean|null):string{
 if(value!==null&&typeof value!=='boolean')throw new OoxmlError('docx-row-header-argument','Row header must be Boolean or null');
 const found=inspect(xml,table,row);if(found.value===value)return xml;
 // on/off also satisfies the SDK's narrower OnOffOnlyType for tblHeader.
 const fragment=value===null?'':`<w:tblHeader xmlns:w="${W}" w:val="${value?'on':'off'}"/>`;
 if(found.marker)return applyEdits(xml,[{start:found.marker.start,end:found.marker.end,value:fragment}]);
 if(!found.pr)return applyEdits(xml,[{start:found.row.openEnd,end:found.row.openEnd,value:`<w:trPr xmlns:w="${W}">${fragment}</w:trPr>`}]);
 const pr=found.pr;if(pr.selfClosing)return applyEdits(xml,[{start:pr.start,end:pr.end,value:xml.slice(pr.start,pr.end).replace(/\/>$/,'>')+fragment+`</${pr.name}>`}]);
 return applyEdits(xml,[{start:pr.closeStart,end:pr.closeStart,value:fragment}]);
}
