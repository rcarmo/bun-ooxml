import {OoxmlError} from '../errors.ts';
import {applyEdits,escapeText,type XmlElement} from '../xml/index.ts';
import {formatRunProperties} from './run-formatting.ts';
import {directParagraphStyle} from './paragraph-style.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS='http://www.w3.org/2000/xmlns/';
const word=(n:XmlElement,name:string)=>n.namespaceURI===W&&n.localName===name;
function fail(message:string):never{throw new OoxmlError('docx-text-unsupported',message);}
function preserveSpace(open:string,name:string):string {
 // Walk complete attributes so namespace values cannot masquerade as xml:space.
 const pattern=/[ \t\r\n]+([^ \t\r\n=]+)[ \t\r\n]*=[ \t\r\n]*(["'])(.*?)\2/ys;
 let cursor=name.length+1;
 while(cursor<open.length){pattern.lastIndex=cursor;const match=pattern.exec(open);if(!match)break;
  if(match[1]==='xml:space'){const end=pattern.lastIndex-1,start=end-match[3]!.length;return open.slice(0,start)+'preserve'+open.slice(end);}
  cursor=pattern.lastIndex;
 }
 return open.replace(/(\/?>)$/,' xml:space="preserve"$1');
}
/** Keep all existing runs/properties; place replacement in the first text leaf. */
export function replaceParagraphText(xml:string,paragraph:XmlElement,text:string):string{
 if(typeof text!=='string'||text.length>1024*1024||/[\t\r\n]/.test(text))throw new OoxmlError('docx-text-argument','Text must be a string of at most 1 Mi UTF-16 units without tabs or line breaks');
 const payload=escapeText(text); // Validate characters before considering a no-op.
 formatRunProperties(xml,paragraph,{});directParagraphStyle(xml,paragraph);
 const runs=paragraph.children.filter(n=>word(n,'r')),leaves=runs.flatMap(r=>r.children.filter(n=>word(n,'t')));
 for(const leaf of leaves){
  for(const key of Object.keys(leaf.attributes))if(leaf.attributeNamespaces[key]!==XMLNS&&key!=='xml:space')fail('Unsupported text-leaf attributes');
  if(leaf.attributes['xml:space']!==undefined&&!['default','preserve'].includes(leaf.attributes['xml:space']))fail('Unsupported xml:space value');
  if(leaf.children.length||!leaf.selfClosing&&xml.slice(leaf.openEnd,leaf.closeStart).includes('<'))fail('Text-leaf lexical content requires preservation, not replacement');
 }
 if(leaves.map(n=>n.text).join('')===text)return xml;
 const preserve=text.startsWith(' ')||text.endsWith(' ');
 if(leaves.length){const edits=leaves.flatMap((leaf,index)=>{
  const value=index===0?payload:'';if(leaf.text===(index===0?text:'')&&!(index===0&&preserve&&leaf.attributes['xml:space']!=='preserve'))return [];
  let open=xml.slice(leaf.start,leaf.openEnd);
  if(index===0&&preserve&&leaf.attributes['xml:space']!=='preserve'){
   open=preserveSpace(open,leaf.name);
  }
  if(leaf.selfClosing)open=open.replace(/\/>$/,'>');
  return [{start:leaf.start,end:leaf.end,value:open+value+`</${leaf.name}>`}];
 });return applyEdits(xml,edits);}
 const t=`<w:t xmlns:w="${W}"${preserve?' xml:space="preserve"':''}>${payload}</w:t>`;
 const owner=runs[0]??paragraph,fragment=runs.length?t:`<w:r xmlns:w="${W}">${t}</w:r>`;
 if(owner.selfClosing)return applyEdits(xml,[{start:owner.start,end:owner.end,value:xml.slice(owner.start,owner.end).replace(/\/>$/,'>')+fragment+`</${owner.name}>`}]);
 return applyEdits(xml,[{start:owner.closeStart,end:owner.closeStart,value:fragment}]);
}
