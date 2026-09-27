import {OoxmlError} from '../errors.ts';
import {applyEdits,attribute,escapeAttribute,type XmlElement} from '../xml/index.ts';
import {formatRunProperties} from './run-formatting.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
const ORDER=['cnfStyle','tcW','gridSpan','hMerge','vMerge','tcBorders','shd','noWrap','tcMar','textDirection','tcFitText','vAlign','hideMark','headers'];
const BORDER_ORDER=['top','left','start','bottom','right','end','insideH','insideV','tl2br','tr2bl'];
export interface CellTopBorder {style:'single';size:number;color:string}
export interface DirectCellProperties {widthTwips:number|null;verticalAlign:'top'|'center'|'bottom'|null;textDirection:'lrTb'|'tbRl'|'btLr'|null;shading:string|null;topBorder:CellTopBorder|null}
export type CellPropertiesPatch=Partial<DirectCellProperties>;
const KEYS=['widthTwips','verticalAlign','textDirection','shading','topBorder'] as const;
const word=(n:XmlElement,name:string)=>n.namespaceURI===W&&n.localName===name;
function fail(message:string):never{throw new OoxmlError('docx-cell-properties-unsupported',message);}
function data(value:unknown,keys:readonly string[]):Record<string,unknown>{
 if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Expected plain cell-property values');
 const out:Record<string,unknown>={};for(const key of Reflect.ownKeys(value)){if(typeof key!=='string'||!keys.includes(key))fail('Unknown cell-property field');const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Cell properties must not be accessors');out[key]=d.value;}return out;
}
function rgb(value:unknown):string{if(typeof value!=='string'||!/^([0-9a-fA-F]{6}|auto)$/.test(value))fail('Expected RGB hex colour or auto');return value;}
function width(value:unknown):number{if(typeof value!=='number'||!Number.isInteger(value)||value<0||value>31680)fail('Cell width must be integer twips from 0 through 31680');return value;}
function border(value:unknown):CellTopBorder{const v=data(value,['style','size','color']);if(v.style!=='single'||typeof v.size!=='number'||!Number.isInteger(v.size)||v.size<2||v.size>96)fail('Top border requires single style and 2 through 96 eighth-points');return {style:'single',size:v.size,color:rgb(v.color)};}
export function normalizeCellProperties(input:CellPropertiesPatch):CellPropertiesPatch{
 const values=data(input,KEYS),out:CellPropertiesPatch={};
 for(const key of KEYS){let v=values[key];if(v===undefined)continue;if(v!==null){
  if(key==='widthTwips')v=width(v);else if(key==='shading')v=rgb(v);else if(key==='topBorder')v=border(v);
  else if(typeof v!=='string'||!(key==='verticalAlign'?['top','center','bottom']:['lrTb','tbRl','btLr']).includes(v))fail('Unsupported alignment or direction');
 }Object.assign(out,{[key]:v});}
 if(!Object.keys(out).length)fail('Expected at least one cell-property field');return out;
}
function noMixed(xml:string,node:XmlElement){if(node.selfClosing)return;let pos=node.openEnd;for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(pos,child.start)))fail('Mixed or lexical cell-property content');pos=child.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(pos,node.closeStart)))fail('Mixed or lexical cell-property content');}
function attrs(node:XmlElement,allowed:readonly string[]){for(const name of Object.keys(node.attributes))if(node.attributeNamespaces[name]!==XMLNS&&!(node.attributeNamespaces[name]===W&&allowed.includes(name.split(':').at(-1)!)))fail('Unknown, themed or misqualified cell-property metadata');}
function ordered(xml:string,node:XmlElement,order:readonly string[]){noMixed(xml,node);attrs(node,[]);let last=-1;const seen=new Set<string>();for(const n of node.children){const rank=order.indexOf(n.localName);if(n.namespaceURI!==W||rank<0||rank<last||seen.has(n.localName))fail('Unknown, revised, duplicate or misordered cell property');seen.add(n.localName);last=rank;}}
function leaf(xml:string,node:XmlElement,allowed:readonly string[]){if(node.children.length)fail('Cell property must be a leaf');noMixed(xml,node);attrs(node,allowed);}
function inspect(xml:string,cell:XmlElement){
 noMixed(xml,cell);const properties=cell.children.filter(n=>word(n,'tcPr'));if(properties.length>1||properties.length&&cell.children[0]!==properties[0])fail('Cell properties must be unique and first');
 const pr=properties[0];if(pr){ordered(xml,pr,ORDER);if(hasRevision(pr))fail('Cell property revisions are unsupported');}
 if(!cell.children.some(n=>word(n,'p')))fail('Cell requires at least one paragraph');
 const nodes=new Map(pr?.children.map(n=>[n.localName,n])??[]);
 if(nodes.has('hMerge')||nodes.has('vMerge')||nodes.has('gridSpan'))fail('Merged or spanned cell properties are unsupported');
 const result:DirectCellProperties={widthTwips:null,verticalAlign:null,textDirection:null,shading:null,topBorder:null};
 for(const n of pr?.children??[]){
  if(n.localName==='tcW'){leaf(xml,n,['w','type']);if(attribute(n,'type',W)!=='dxa'||!/^\d+$/.test(attribute(n,'w',W)??''))fail('Only explicit twip cell widths are supported');result.widthTwips=width(Number(attribute(n,'w',W)));}
  if(n.localName==='vAlign'||n.localName==='textDirection'){leaf(xml,n,['val']);const v=attribute(n,'val',W);if(!v||!(n.localName==='vAlign'?['top','center','bottom']:['lrTb','tbRl','btLr']).includes(v))fail('Unsupported direct cell alignment/direction');if(n.localName==='vAlign')result.verticalAlign=v as DirectCellProperties['verticalAlign'];else result.textDirection=v as DirectCellProperties['textDirection'];}
  if(n.localName==='shd'){leaf(xml,n,['val','color','fill']);if(attribute(n,'val',W)!=='clear'||!['auto',undefined].includes(attribute(n,'color',W)))fail('Only clear background shading with automatic foreground is supported');result.shading=rgb(attribute(n,'fill',W));}
  if(n.localName==='tcBorders'){ordered(xml,n,BORDER_ORDER);for(const b of n.children){if(b.children.length)fail('Nested border metadata');noMixed(xml,b);}const top=n.children.find(c=>word(c,'top'));if(top){leaf(xml,top,['val','sz','color']);const sz=attribute(top,'sz',W);if(!sz||!/^\d+$/.test(sz))fail('Missing top border size');result.topBorder=border({style:attribute(top,'val',W),size:Number(sz),color:attribute(top,'color',W)});}}
 }
 return {pr,nodes,result};
}
export function readCellProperties(xml:string,cell:XmlElement):DirectCellProperties{return inspect(xml,cell).result;}
export function editCellProperties(xml:string,cell:XmlElement,patch:CellPropertiesPatch):string{
 const {pr,nodes,result}=inspect(xml,cell);
 for(const child of cell.children){if(word(child,'tcPr'))continue;if(!word(child,'p'))fail('Cell properties require plain paragraphs');formatRunProperties(xml,child,{});}
 // Reject changes/revision wrappers on the direct cell/row/table path.
 for(let parent=cell.parent;parent&&parent.localName!=='body';parent=parent.parent){
  noMixed(xml,parent);const allowed=word(parent,'tr')?['trPr','tc']:word(parent,'tbl')?['tblPr','tblGrid','tr']:[];
  const seen=new Set<string>();let last=-1;for(const n of parent.children){const rank=allowed.indexOf(n.localName);if(n.namespaceURI!==W||rank<0||rank<last)fail('Unsupported or misordered table owner topology');last=rank;if(n.localName!=='tc'&&n.localName!=='tr'){if(seen.has(n.localName))fail('Duplicate table metadata');seen.add(n.localName);if(hasRevision(n))fail('Table property revisions are unsupported');}}
  if(word(parent,'tbl')&&!seen.has('tblGrid'))fail('Formatting requires an explicit table grid');
 }
 const edits:{start:number;end:number;value:string}[]=[],additions:{name:string;value:string}[]=[];
 const replace=(name:string,value:string)=>{const n=nodes.get(name);if(n)edits.push({start:n.start,end:n.end,value});else if(value)additions.push({name,value});};
 const val=(name:string,value:string)=>`<w:${name} xmlns:w="${W}" w:val="${escapeAttribute(value)}"/>`;
 for(const key of KEYS){const value=patch[key];if(value===undefined||JSON.stringify(value)===JSON.stringify(result[key]))continue;
  if(key==='widthTwips')replace('tcW',value===null?'':`<w:tcW xmlns:w="${W}" w:w="${value}" w:type="dxa"/>`);
  else if(key==='shading')replace('shd',value===null?'':`<w:shd xmlns:w="${W}" w:val="clear" w:fill="${value}"/>`);
  else if(key==='verticalAlign'||key==='textDirection')replace(key==='verticalAlign'?'vAlign':'textDirection',value===null?'':val(key==='verticalAlign'?'vAlign':'textDirection',String(value)));
  else {
   const b=value as CellTopBorder|null,serialized=b?`<w:top xmlns:w="${W}" w:val="single" w:sz="${b.size}" w:color="${b.color}"/>`:'',borders=nodes.get('tcBorders'),top=borders?.children.find(n=>word(n,'top'));
   if(top)edits.push({start:top.start,end:top.end,value:serialized});
   else if(borders&&serialized){if(borders.selfClosing)edits.push({start:borders.start,end:borders.end,value:xml.slice(borders.start,borders.end).replace(/\/>$/,'>')+serialized+`</${borders.name}>`});else edits.push({start:borders.openEnd,end:borders.openEnd,value:serialized});}
   else if(serialized)additions.push({name:'tcBorders',value:`<w:tcBorders xmlns:w="${W}">${serialized}</w:tcBorders>`});
  }
 }
 additions.sort((a,b)=>ORDER.indexOf(a.name)-ORDER.indexOf(b.name));
 if(additions.length){if(!pr||pr.selfClosing){const content=additions.map(a=>a.value).join('');if(pr)edits.push({start:pr.start,end:pr.end,value:xml.slice(pr.start,pr.end).replace(/\/>$/,'>')+content+`</${pr.name}>`});else edits.push({start:cell.openEnd,end:cell.openEnd,value:`<w:tcPr xmlns:w="${W}">${content}</w:tcPr>`});}
 else{const grouped=new Map<number,string>();for(const a of additions){const next=pr.children.find(n=>ORDER.indexOf(n.localName)>ORDER.indexOf(a.name)),at=next?.start??pr.closeStart;grouped.set(at,(grouped.get(at)??'')+a.value);}for(const[start,value]of grouped)edits.push({start,end:start,value});}}
 return edits.length?applyEdits(xml,edits):xml;
}
function hasRevision(node:XmlElement):boolean{return node.namespaceURI===W&&node.localName.endsWith('Change')||node.children.some(hasRevision);}
