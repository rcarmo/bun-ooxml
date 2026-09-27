import {OoxmlError} from '../errors.ts';
import {applyEdits,escapeText,parseXml,type XmlElement} from '../xml/index.ts';
import {formatRunProperties,type DirectRunPatch} from './run-formatting.ts';
import {directParagraphStyle} from './paragraph-style.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
/** Build one independent plain-text run, without borrowing preceding formatting. */
export function appendParagraphRun(xml:string,paragraph:XmlElement,text:string,formatting:DirectRunPatch):string {
 if(typeof text!=='string'||text.length>1024*1024||/[\t\r\n]/.test(text))throw new OoxmlError('docx-run-argument','Run text must be a string of at most 1 Mi UTF-16 units without tabs or line breaks');
 // Both existing paragraph/property topology and new run formatting preflight
 // before any package changes. No source nodes are mutated.
 formatRunProperties(xml,paragraph,{});
 directParagraphStyle(xml,paragraph);
 const payload=escapeText(text),space=/^[ \t\r\n]|[ \t\r\n]$/.test(text)?' xml:space="preserve"':'';
 const sample=`<w:p xmlns:w="${W}"><w:r xmlns:w="${W}"><w:t${space}>${payload}</w:t></w:r></w:p>`;
 const rendered=formatRunProperties(sample,parseXml(sample).root,formatting).xml;
 const run=parseXml(rendered).root.children[0]!;
 const fragment=rendered.slice(run.start,run.end);
 if(paragraph.selfClosing){
  const original=xml.slice(paragraph.start,paragraph.end),opening=original.replace(/\/>$/,'>');
  return applyEdits(xml,[{start:paragraph.start,end:paragraph.end,value:opening+fragment+`</${paragraph.name}>`}]);
 }
 return applyEdits(xml,[{start:paragraph.closeStart,end:paragraph.closeStart,value:fragment}]);
}
