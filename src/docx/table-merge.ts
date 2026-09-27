import {OoxmlError} from '../errors.ts';
import {applyEdits,attribute,type XmlElement} from '../xml/index.ts';
import {readRowHeader} from './row-header.ts';
import {readCellProperties} from './cell-properties.ts';
import {formatRunProperties} from './run-formatting.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
const word=(n:XmlElement,name:string)=>n.namespaceURI===W&&n.localName===name;
function fail(message:string):never{throw new OoxmlError('docx-table-merge-unsupported',message);}
function attrs(n:XmlElement,names:string[]=[]){for(const key of Object.keys(n.attributes))if(n.attributeNamespaces[key]!==XMLNS&&!(n.attributeNamespaces[key]===W&&names.includes(key.split(':').at(-1)!)))fail('Decorated merge input requires review');}
function gaps(xml:string,n:XmlElement){if(n.selfClosing)return;let cursor=n.openEnd;for(const c of n.children){if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,c.start)))fail('Mixed or lexical merge content');cursor=c.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,n.closeStart)))fail('Mixed or lexical merge content');}
/** Merge one row's plain, unmerged cells. Never discard absorbed content or formatting. */
export function mergeRowCells(xml:string,table:XmlElement,row:number,first:number,last:number):{xml:string;removedParagraphStarts:number[]}{
 const rows=table.children.filter(n=>word(n,'tr')),grids=table.children.filter(n=>word(n,'tblGrid'));
 if(![row,first,last].every(Number.isSafeInteger)||row<0||row>=rows.length||first<0||last<=first)throw new RangeError('Merge requires a row and at least two increasing columns');
 if(grids.length!==1)fail('One explicit grid is required');const grid=grids[0]!;attrs(grid);gaps(xml,grid);
 if(grid.children.length<1||grid.children.length>100||rows.length>100)fail('Merge supports bounded tables up to 100 by 100');
 if(last>=grid.children.length)throw new RangeError('Merge column is outside the grid');
 const widths=grid.children.map(n=>{attrs(n,['w']);gaps(xml,n);const v=attribute(n,'w',W);if(!word(n,'gridCol')||n.children.length||!v||!/^\d+$/.test(v)||Number(v)>31680)fail('Grid widths must be explicit bounded twips');return Number(v);});
 // Validate the whole table before editing. Existing merges, revisions, nested
 // cells and ambiguous row/property topology remain outside this authoring API.
 for(const [i,r]of rows.entries()){
  readRowHeader(xml,table,i);const cells=r.children.filter(n=>word(n,'tc'));if(cells.length!==widths.length)fail('Unmerged rows must match the grid');
  for(const [col,cell]of cells.entries()){
   if(cell.children.some(n=>!word(n,'tcPr')&&!word(n,'p')))fail('Only direct plain cell paragraphs are supported');
   if(readCellProperties(xml,cell).widthTwips!==widths[col])fail('Cell width must match its grid column');
   for(const p of cell.children.filter(n=>word(n,'p'))){if(p.children.some(pr=>word(pr,'pPr')&&pr.children.some(c=>word(c,'sectPr'))))fail('Section-bearing cell paragraphs require section-aware editing');formatRunProperties(xml,p,{});}
  }
 }
 const cells=rows[row]!.children.filter(n=>word(n,'tc')),owner=cells[first]!,absorbed=cells.slice(first+1,last+1);
 for(const cell of absorbed){
  attrs(cell);gaps(xml,cell);const prs=cell.children.filter(n=>word(n,'tcPr')),paragraphs=cell.children.filter(n=>word(n,'p'));
  if(prs.length!==1||paragraphs.length!==1||cell.children.length!==2)fail('Absorbed cells must have only a width and one empty paragraph');
  const pr=prs[0]!;attrs(pr);gaps(xml,pr);if(pr.children.length!==1||!word(pr.children[0]!,'tcW'))fail('Absorbed cell formatting must not be discarded');
  const p=paragraphs[0]!;attrs(p);gaps(xml,p);if(p.children.length)fail('Absorbed cell paragraphs must be structurally empty');
 }
 const width=widths.slice(first,last+1).reduce((a,b)=>a+b,0);if(width>31680)fail('Merged cell width exceeds the bounded twip range');
 const pr=owner.children.find(n=>word(n,'tcPr'))!,w=pr.children.find(n=>word(n,'tcW'))!;
 const replacement=`<w:tcW xmlns:w="${W}" w:w="${width}" w:type="dxa"/><w:gridSpan xmlns:w="${W}" w:val="${last-first+1}"/>`;
 const edits=[{start:w.start,end:w.end,value:replacement},...absorbed.map(c=>({start:c.start,end:c.end,value:''}))];
 return {xml:applyEdits(xml,edits),removedParagraphStarts:absorbed.flatMap(c=>c.children.filter(n=>word(n,'p')).map(n=>n.start))};
}
