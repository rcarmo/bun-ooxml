import {OoxmlError} from '../errors.ts';
import {applyEdits,attribute,escapeAttribute,type XmlElement} from '../xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS='http://www.w3.org/2000/xmlns/';
export const PARAGRAPH_PROPERTY_ORDER=['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl','numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap','overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid','spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection','textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr'];
function fail(message:string):never{throw new OoxmlError('docx-style-unsupported',message);}
const word=(n:XmlElement,name:string)=>n.namespaceURI===W&&n.localName===name;
function noMixed(n:XmlElement,xml:string){if(n.selfClosing)return;let cursor=n.openEnd;for(const c of n.children){if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,c.start)))fail('Mixed paragraph property content is unsupported');cursor=c.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,n.closeStart)))fail('Mixed paragraph property content is unsupported');}
function revised(n:XmlElement):boolean{return n.namespaceURI===W&&n.localName.endsWith('Change')||n.children.some(revised);}
function inspect(xml:string,p:XmlElement){
 const properties=p.children.filter(c=>word(c,'pPr'));if(properties.length>1||properties.length&&p.children[0]!==properties[0])fail('Paragraph properties must be unique and first');
 const pr=properties[0];if(!pr)return {};
 noMixed(pr,xml);let last=-1;const seen=new Set<string>();
 for(const n of pr.children){const rank=PARAGRAPH_PROPERTY_ORDER.indexOf(n.localName);if(n.namespaceURI!==W||rank<0||rank<last||seen.has(n.localName)||revised(n))fail('Unknown, duplicate, misplaced or revised paragraph properties');seen.add(n.localName);last=rank;}
 const node=pr.children.find(c=>word(c,'pStyle'));if(!node)return {pr};
 if(node.children.length||!node.selfClosing&&xml.slice(node.openEnd,node.closeStart).trim())fail('Style override cannot contain content');
 for(const k of Object.keys(node.attributes))if(node.attributeNamespaces[k]!==XMLNS&&!(k.split(':').at(-1)==='val'&&node.attributeNamespaces[k]===W))fail('Style ID attribute must use the Word namespace');
 const id=attribute(node,'val',W);if(!id||!id.trim())fail('Empty or missing paragraph style ID');return {pr,node,id};
}
export function directParagraphStyle(xml:string,p:XmlElement):string|undefined{return inspect(xml,p).id;}
export function replaceParagraphStyle(xml:string,p:XmlElement,id:string|null):string{
 const {pr,node,id:current}=inspect(xml,p);
 noMixed(p,xml);
 for(const r of p.children){if(word(r,'pPr'))continue;if(!word(r,'r'))fail('Style mutation requires plain direct text runs');noMixed(r,xml);for(const c of r.children)if(!word(c,'rPr')&&!word(c,'t')||word(c,'t')&&c.children.length||revised(c))fail('Unsupported or revised run content');}
 if(current===(id??undefined))return xml;
 if(id===null)return node?applyEdits(xml,[{start:node.start,end:node.end,value:''}]):xml;
 const style=`<w:pStyle xmlns:w="${W}" w:val="${escapeAttribute(id)}"/>`;
 if(node)return applyEdits(xml,[{start:node.start,end:node.end,value:style}]);
 if(pr){
  if(pr.selfClosing){const open=xml.slice(pr.start,pr.end).replace(/\/\s*>$/,'>');return applyEdits(xml,[{start:pr.start,end:pr.end,value:open+style+`</${pr.name}>`}]);}
  return applyEdits(xml,[{start:pr.openEnd,end:pr.openEnd,value:style}]);
 }
 const properties=`<w:pPr xmlns:w="${W}">${style}</w:pPr>`;
 if(p.selfClosing){const open=xml.slice(p.start,p.end).replace(/\/\s*>$/,'>');return applyEdits(xml,[{start:p.start,end:p.end,value:open+properties+`</${p.name}>`}]);}
 return applyEdits(xml,[{start:p.openEnd,end:p.openEnd,value:properties}]);
}
