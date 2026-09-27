import {OoxmlError} from '../errors.ts';
import {applyEdits,type XmlDocument,type XmlElement} from '../xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const word=(n:XmlElement,name:string)=>n.namespaceURI===W&&n.localName===name;
function fail(message:string):never{throw new OoxmlError('docx-body-unsupported',message);}
/** Body indexes count direct paragraphs and tables, never final sectPr or cell paragraphs. */
export function bodyInsertionTarget(xml:string,document:XmlDocument,index:number):{body:XmlElement;start:number}{
 if(!word(document.root,'document'))fail('Expected a Word document root');
 const bodies=document.root.children.filter(n=>word(n,'body'));if(bodies.length!==1)fail('Expected one body');const body=bodies[0]!;
 let cursor=body.openEnd,seenSection=false;const blocks:XmlElement[]=[];
 for(const child of body.children){
  if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,child.start)))fail('Mixed or lexical body content');cursor=child.end;
  if(word(child,'sectPr')){if(seenSection||child!==body.children.at(-1))fail('Section properties must be unique and last');seenSection=true;}
  else if(word(child,'p')||word(child,'tbl'))blocks.push(child);
  else fail('Body indexes require direct paragraphs and tables');
 }
 if(!body.selfClosing&&!/^[ \t\r\n]*$/.test(xml.slice(cursor,body.closeStart)))fail('Mixed or lexical body content');
 if(!Number.isInteger(index)||index<0||index>blocks.length)throw new RangeError('DOCX body index must be between zero and the number of top-level paragraphs and tables');
 return {body,start:blocks[index]?.start??body.children.find(n=>word(n,'sectPr'))?.start??body.closeStart};
}
export function insertBodyParagraph(xml:string,target:{body:XmlElement;start:number},fragment:string):string{
 const {body,start}=target;
 if(body.selfClosing)return applyEdits(xml,[{start:body.start,end:body.end,value:xml.slice(body.start,body.end).replace(/\/>$/,'>')+fragment+`</${body.name}>`}]);
 return applyEdits(xml,[{start,end:start,value:fragment}]);
}
