import {OoxmlError} from '../errors.ts';
import {applyEdits,attribute,type XmlElement} from '../xml/index.ts';
import {readRowHeader} from './row-header.ts';
import {replaceParagraphText} from './paragraph-text.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
const word=(node:XmlElement,name:string)=>node.namespaceURI===W&&node.localName===name;
function fail(message:string):never{throw new OoxmlError('docx-table-rows-unsupported',message);}
function attrs(node:XmlElement,allowed:string[]){for(const name of Object.keys(node.attributes))if(node.attributeNamespaces[name]!==XMLNS&&!(node.attributeNamespaces[name]===W&&allowed.includes(name.split(':').at(-1)!)))fail('Unsupported row or grid attributes');}
function gaps(xml:string,node:XmlElement){if(node.selfClosing)return;let pos=node.openEnd;for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(pos,child.start)))fail('Lexical or mixed grid content');pos=child.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(pos,node.closeStart)))fail('Lexical or mixed grid content');}
/** Validate every row before changing the table topology; no merged or nested cells. */
export function editTableRow(xml:string,table:XmlElement,index:number,operation:'insert'|'delete'):{xml:string;rows:XmlElement[];removed?:XmlElement;start:number}{
 const rows=table.children.filter(n=>word(n,'tr'));
 if(!Number.isInteger(index)||index<0||index>(operation==='insert'?rows.length:rows.length-1))throw new RangeError('DOCX row index is out of range');
 if(rows.length<1||rows.length>100||operation==='insert'&&rows.length===100||operation==='delete'&&rows.length===1)throw new RangeError('Row edits require between one and 100 resulting rows');
 const grids=table.children.filter(n=>word(n,'tblGrid'));if(grids.length!==1)fail('One explicit table grid is required');const grid=grids[0]!;attrs(grid,[]);gaps(xml,grid);
 if(grid.children.length<1||grid.children.length>100)fail('Grid must contain one through 100 columns');
 const widths=grid.children.map(col=>{if(!word(col,'gridCol')||col.children.length)fail('Only plain grid columns are supported');attrs(col,['w']);gaps(xml,col);const width=attribute(col,'w',W);if(width===undefined||!/^\d+$/.test(width)||Number(width)>31680)fail('Explicit grid widths must be integer twips from zero through 31680');return Number(width);});
 for(const [i,row]of rows.entries()){
  readRowHeader(xml,table,i);attrs(row,['rsidR','rsidRPr','rsidTr']);const cells=row.children.filter(n=>word(n,'tc'));if(cells.length!==widths.length)fail('Rows must match the explicit grid');
  for(const cell of cells){attrs(cell,[]);for(const p of cell.children)if(word(p,'p')){
   attrs(p,['rsidR','rsidRPr','rsidP','rsidRDefault']);
   if(p.children.some(pr=>word(pr,'pPr')&&pr.children.some(n=>word(n,'sectPr'))))fail('Section-bearing cell paragraphs require section-aware editing');
   const runs=p.children.filter(n=>word(n,'r'));for(const run of runs)attrs(run,['rsidR','rsidRPr','rsidDel']);
   const text=runs.flatMap(r=>r.children.filter(n=>word(n,'t'))).map(n=>n.text).join('');replaceParagraphText(xml,p,text);
  }}
 }
 const removed=operation==='delete'?rows[index]:undefined;
 if(removed)return {xml:applyEdits(xml,[{start:removed.start,end:removed.end,value:''}]),rows,removed,start:removed.start};
 const start=rows[index]?.start??table.closeStart;
 const fragment=`<w:tr xmlns:w="${W}">${widths.map(width=>`<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`).join('')}</w:tr>`;
 return {xml:applyEdits(xml,[{start,end:start,value:fragment}]),rows,start};
}
