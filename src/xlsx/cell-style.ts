import {OoxmlError} from '../errors.ts';
import {OpcPackage,getContentType} from '../opc/index.ts';
import {parseXml,elements,applyEdits,type XmlElement} from '../xml/index.ts';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles';
const MIME='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml';
const fail=(message:string):never=>{throw new OoxmlError('xlsx-style-invalid',message);};
const is=(node:XmlElement,local:string)=>node.namespaceURI===S&&node.localName===local;
function index(value:string,label:string):number{
 if(!/^\d+$/.test(value)||!Number.isSafeInteger(Number(value))||Number(value)>4294967295)fail(`Invalid ${label}`);return Number(value);
}
function collection(root:XmlElement,name:string,child:string):XmlElement[]{
 const matches=root.children.filter(n=>n.localName===name);if(matches.length>1||matches.some(n=>!is(n,name)))fail(`Ambiguous ${name} collection`);
 const container=matches[0];if(!container)return [];
 if(container.children.some(n=>!is(n,child)))fail(`Unsupported ${name} collection child`);
 if(container.attributes.count!==undefined&&index(container.attributes.count,`${name} count`)!==container.children.length)fail(`Incorrect ${name} count`);
 return container.children;
}
/** Validate the selected cellXf and its base dependencies without changing definitions. */
function validateStyle(pkg:OpcPackage,main:string,selected:number){
 const links=pkg.relationships(main).filter(r=>r.type===REL);
 if(links.length!==1||links[0]!.external||!links[0]!.resolved)fail('Expected one internal styles relationship');
 const part=links[0]!.resolved!;if(getContentType(pkg,part)!==MIME)fail('Wrong styles content type');
 const root=parseXml(pkg.text(part)).root;if(!is(root,'styleSheet'))fail('Wrong styles root');
 const cellXfs=collection(root,'cellXfs','xf'),baseXfs=collection(root,'cellStyleXfs','xf'),fonts=collection(root,'fonts','font'),fills=collection(root,'fills','fill'),borders=collection(root,'borders','border'),formats=collection(root,'numFmts','numFmt');
 const custom=new Set<number>();for(const node of formats){const id=index(node.attributes.numFmtId??'','number format ID');if(id<164||custom.has(id)||!node.attributes.formatCode)fail('Invalid custom number format');custom.add(id);}
 const validate=(xf:XmlElement)=>{
  for(const [key,entries]of [['fontId',fonts],['fillId',fills],['borderId',borders]] as const){if(index(xf.attributes[key]??'0',key)>=entries.length)fail(`Missing ${key} dependency`);}
  const num=index(xf.attributes.numFmtId??'0','numFmtId');if(num>=164&&!custom.has(num))fail('Missing custom number format');
  for(const key of Object.keys(xf.attributes))if(['fontId','fillId','borderId','numFmtId','xfId'].includes(key.split(':').at(-1)!)&&key.includes(':'))fail('Style references must be unqualified');
 };
 const xf=cellXfs[selected];if(!xf)fail('Missing cellXf index');validate(xf!);
 if(xf!.attributes.xfId!==undefined){const base=baseXfs[index(xf!.attributes.xfId,'xfId')];if(!base)fail('Missing base style');validate(base!);if(base!.attributes.xfId!==undefined)fail('Base style xfId chains are unsupported');}
}

export function selectCellStyle(pkg:OpcPackage,main:string,xml:string,cell:XmlElement,selected:number|null):string {
 if(selected!==null&&(!Number.isSafeInteger(selected)||selected<0||selected>4294967295))fail('Style index must be a nonnegative integer or null');
 const root=parseXml(xml).root;
 if(!is(root,'worksheet')||elements(root,'sheetProtection',S).length||elements(parseXml(pkg.text(main)),'workbookProtection',S).length)fail('Worksheet/workbook protection refuses style changes');
 // The index model can read more topologies than this mutation surface admits.
 const data=root.children.filter(n=>is(n,'sheetData'));if(data.length!==1)fail('Expected one sheetData collection');
 const row=cell.parent,sourceData=row?.parent;
 if(!row||!is(row,'row')||!sourceData||!is(sourceData,'sheetData')||sourceData.start!==data[0]!.start)fail('Cell is not a direct worksheet row cell');
 const ref=cell.attributes.r??'',rowNumber=Number(ref.match(/\d+$/)?.[0]);
 if(!/^\d+$/.test(row!.attributes.r??'')||Number(row!.attributes.r)!==rowNumber||data[0]!.children.filter(n=>is(n,'row')&&Number(n.attributes.r)===rowNumber).length!==1)fail('Cell row is ambiguous');
 if(Object.keys(cell.attributes).some(k=>k.includes(':')&&k.split(':').at(-1)==='s'))fail('Cell style index must be unqualified');
 const current=cell.attributes.s===undefined?undefined:index(cell.attributes.s,'current style index');
 if(selected!==null)validateStyle(pkg,main,selected);
 if(selected===null?current===undefined:current===selected)return xml;
 const open=xml.slice(cell.start,cell.openEnd);let style:RegExpMatchArray|undefined;
 for(const match of open.matchAll(/\s+([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/g))if(match[1]==='s')style=match;
 if(style){const at=cell.start+style.index!;return applyEdits(xml,[{start:at,end:at+style[0].length,value:selected===null?'':style[0].replace(/(["'])[\s\S]*\1$/,`"${selected}"`)}]);}
 const tail=open.match(/\/?\s*>$/)!;const at=cell.openEnd-tail[0].length;
 return applyEdits(xml,[{start:at,end:at,value:` s="${selected}"`}]);
}
