import {OoxmlError} from '../errors.ts';
import {attribute,type XmlDocument,type XmlElement} from '../xml/index.ts';
import {bodyInsertionTarget} from './body-insertion.ts';
import {formatRunProperties} from './run-formatting.ts';
import {readParagraphProperties} from './paragraph-properties.ts';
import {directParagraphStyle} from './paragraph-style.ts';
import {readRowHeader} from './row-header.ts';
import {readCellProperties} from './cell-properties.ts';
import {literalPlaceholders} from './body-map.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
const word=(n:XmlElement,s:string)=>n.namespaceURI===W&&n.localName===s;
export interface TemplateParagraph {readonly index:number;readonly bodyIndex:number;readonly text:string;readonly styleId:string|null;readonly outlineLevel:number|null;readonly table?:Readonly<{index:number;row:number;column:number;paragraph:number}>}
export interface TemplateSection {readonly paragraphIndex:number;readonly bodyIndex:number;readonly text:string;readonly level:number}
export interface TemplateTable {readonly index:number;readonly bodyIndex:number;readonly rows:number;readonly columns:number;readonly cells:readonly (readonly string[])[]}
export interface TemplatePlaceholder {readonly paragraphIndex:number;readonly bodyIndex:number;readonly start:number;readonly end:number;readonly text:string;readonly name:string}
export interface TemplateInspection {
 readonly scope:'main-body-and-tables';
 readonly counts:Readonly<{paragraphs:number;sections:number;tables:number;placeholders:number}>;
 readonly paragraphs:readonly TemplateParagraph[];readonly sections:readonly TemplateSection[];readonly tables:readonly TemplateTable[];readonly placeholders:readonly TemplatePlaceholder[];
}
function fail(message:string):never{throw new OoxmlError('docx-template-unsupported',message);}
function limit(message:string):never{throw new OoxmlError('docx-template-limit',message);}
/** Explicit source inventory only: no semantic classifications or cache state. */
export function inspectTemplateXml(xml:string,document:XmlDocument):TemplateInspection {
 if(!word(document.root,'document'))fail('Template inspection requires a Transitional Word document');
 const {body}=bodyInsertionTarget(xml,document,0),paragraphs:TemplateParagraph[]=[],sections:TemplateSection[]=[],tables:TemplateTable[]=[],placeholders:TemplatePlaceholder[]=[];
 let bodyIndex=0;
 const whitespace=(n:XmlElement)=>{if(n.selfClosing)return;let at=n.openEnd;for(const c of n.children){if(!/^[ \t\r\n]*$/.test(xml.slice(at,c.start)))fail('Mixed or lexical grid content');at=c.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(at,n.closeStart)))fail('Mixed or lexical grid content');};
 const paragraph=(p:XmlElement,table?:TemplateParagraph['table']):string=>{
  if(paragraphs.length===10000)limit('At most 10000 template paragraphs are supported');
  formatRunProperties(xml,p,{});
  const text=p.children.filter(n=>word(n,'r')).flatMap(r=>r.children.filter(n=>word(n,'t'))).map(n=>n.text).join('');
  const outlineLevel=readParagraphProperties(xml,p).outlineLevel,styleId=directParagraphStyle(xml,p)??null,index=paragraphs.length;
  paragraphs.push(Object.freeze({index,bodyIndex,text,styleId,outlineLevel,...(table?{table:Object.freeze(table)}:{})}));
  if(!table&&outlineLevel!==null&&outlineLevel<9)sections.push(Object.freeze({paragraphIndex:index,bodyIndex,text,level:outlineLevel+1}));
  for(const occurrence of literalPlaceholders(text)){if(placeholders.length===10000)limit('At most 10000 template placeholders are supported');placeholders.push(Object.freeze({paragraphIndex:index,bodyIndex,...occurrence}));}
  return text;
 };
 for(const node of body.children){
  if(word(node,'sectPr'))continue;
  if(word(node,'p'))paragraph(node);
  else if(word(node,'tbl')){
   if(tables.length===1000)limit('At most 1000 template tables are supported');
   const grids=node.children.filter(n=>word(n,'tblGrid')),rows=node.children.filter(n=>word(n,'tr'));
   if(grids.length!==1||rows.length===0)fail('Template tables require one explicit grid and at least one row');
   const grid=grids[0]!;
   if(rows.length>100||grid.children.length===0||grid.children.length>100)limit('Template tables support one through 100 rows and columns');
   whitespace(grid);for(const key of Object.keys(grid.attributes))if(grid.attributeNamespaces[key]!==XMLNS)fail('Decorated table grid is unsupported');
   const widths:number[]=[];
   for(const col of grid.children){whitespace(col);for(const key of Object.keys(col.attributes))if(col.attributeNamespaces[key]!==XMLNS&&!(col.attributeNamespaces[key]===W&&key.split(':').at(-1)==='w'))fail('Decorated grid width is unsupported');const width=attribute(col,'w',W);if(!word(col,'gridCol')||col.children.length||width===undefined||!/^\d+$/.test(width)||Number(width)>31680)fail('Template grid widths must be explicit bounded twips');widths.push(Number(width));}
   const index=tables.length,matrix:ReadonlyArray<string>[]=[];
   for(const [row,r]of rows.entries()){
    readRowHeader(xml,node,row);
    for(const offset of r.children.find(n=>word(n,'trPr'))?.children.filter(n=>word(n,'gridBefore')||word(n,'gridAfter'))??[]){const value=attribute(offset,'val',W);if(value===undefined||!/^0+$/.test(value))fail('Template rows cannot omit grid columns');}
    const cells=r.children.filter(n=>word(n,'tc'));if(cells.length!==grid.children.length)fail('Template table rows must match the grid');
    const values:string[]=[];
    for(const [column,c]of cells.entries()){
     if(c.children.some(n=>!word(n,'tcPr')&&!word(n,'p')))fail('Only plain template cell paragraphs are supported');
     if(readCellProperties(xml,c).widthTwips!==widths[column])fail('Template cell widths must match their grid columns');
     const ps=c.children.filter(n=>word(n,'p'));if(!ps.length)fail('Template cells require paragraphs');
     values.push(ps.map((p,i)=>paragraph(p,{index,row,column,paragraph:i})).join('\n'));
    }
    matrix.push(Object.freeze(values));
   }
   tables.push(Object.freeze({index,bodyIndex,rows:rows.length,columns:grid.children.length,cells:Object.freeze(matrix)}));
  }else fail('Unsupported template body block');
  bodyIndex++;
 }
 return Object.freeze({scope:'main-body-and-tables',counts:Object.freeze({paragraphs:paragraphs.length,sections:sections.length,tables:tables.length,placeholders:placeholders.length}),paragraphs:Object.freeze(paragraphs),sections:Object.freeze(sections),tables:Object.freeze(tables),placeholders:Object.freeze(placeholders)});
}
