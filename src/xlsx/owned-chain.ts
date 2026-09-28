import {OoxmlError} from '../errors.ts';
import {getContentType} from '../opc/content-types.ts';
import {relationshipPath,type OpcPackage} from '../opc/package.ts';
import {analyzeFormulaReferences} from './formula.ts';
import {parseA1Range} from './range.ts';
import {attribute,elements,parseXml,type XmlDocument,type XmlElement} from '../xml/index.ts';

const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CHAIN_REL=`${OFFICE}/calcChain`;
const SHEET_REL=`${OFFICE}/worksheet`;
const CHAIN_TYPE='application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
const CT='http://schemas.openxmlformats.org/package/2006/content-types';
type Sheet={name:string;part:string;document:XmlDocument;refs:ReadonlySet<string>};
export type OwnedChainPlan={part:string;relationshipId:string;affected:ReadonlySet<string>};
const refuse=(message:string):never=>{throw new OoxmlError('xlsx-calculation-chain-unsupported',message);};
const children=(node:XmlElement,name:string,ns:string)=>node.children.filter(e=>e.localName===name&&e.namespaceURI===ns);
const exact=(node:XmlElement,keys:string[])=>Object.keys(node.attributes).every(k=>(keys.includes(k)&&node.attributeNamespaces[k]==='')||((k==='xmlns'||k.startsWith('xmlns:'))&&node.attributeNamespaces[k]==='http://www.w3.org/2000/xmlns/'));
function cellReference(text:string):boolean{try{const r=parseA1Range(text);return r.axis==='cell'&&!r.sheet&&r.first.row===r.last.row&&r.first.column===r.last.column&&text===text.toUpperCase()&&!text.includes('$')&&!text.includes(':');}catch{return false;}}

/** The opt-in path requires a complete, static dependency graph. Chain entries
 * are order metadata only; they never supply dependency information. */
export function inspectOwnedChain(pkg:OpcPackage,workbookPart:string,sheets:readonly Sheet[],editedSheet:string,editedRef:string):OwnedChainPlan{
 const rels=pkg.relationships(workbookPart);
 const chains=rels.filter(r=>r.type.endsWith('/calcChain'));
 if(chains.length!==1||chains[0]!.type!==CHAIN_REL||chains[0]!.external||!chains[0]!.resolved||chains[0]!.target.includes('#'))refuse('Expected one internal workbook-owned calculation chain');
 const chain=chains[0]!,part=chain.resolved!;
 if(part==='xl/calcChain.xml'||!part.startsWith('xl/')||!part.endsWith('.xml')||sheets.some(s=>s.part===part)||!pkg.get(part))refuse('Unsupported calculation-chain part');
 if(pkg.get(relationshipPath(part)))refuse('Calculation chain has an outgoing relationship registry');
 const names=pkg.names();
 if(names.some(n=>n==='xl/calcChain.xml'||n!==part&&getContentType(pkg,n)===CHAIN_TYPE))refuse('Another calculation-chain part is present');
 const types=parseXml(pkg.text('[Content_Types].xml'));
 const overrides=elements(types.root,'Override',CT).filter(e=>e.attributes.PartName===`/${part}`||e.attributes.ContentType===CHAIN_TYPE);
 if(overrides.length!==1||overrides[0]!.attributes.PartName!==`/${part}`||overrides[0]!.attributes.ContentType!==CHAIN_TYPE||getContentType(pkg,part)!==CHAIN_TYPE)refuse('Missing or ambiguous calculation-chain content type');
 for(const owner of ['',...names.filter(n=>n!=='[Content_Types].xml'&&!n.endsWith('.rels'))]){
  for(const r of pkg.relationships(owner))if((r.type.endsWith('/calcChain')||!r.external&&r.resolved===part)&&!(owner===workbookPart&&r.id===chain.id))refuse('Calculation chain has another owner');
 }
 const workbook=parseXml(pkg.text(workbookPart));
 if(workbook.root.localName!=='workbook'||workbook.root.namespaceURI!==S)refuse('Invalid workbook root');
 const lists=children(workbook.root,'sheets',S),declared=lists.length===1?children(lists[0]!,'sheet',S):[];
 if(lists.length!==1||declared.length!==sheets.length||lists[0]!.children.length!==sheets.length||new Set(sheets.map(s=>s.part)).size!==sheets.length)refuse('Incomplete worksheet ownership');
 if(sheets.length>128)refuse('Too many worksheets for dependency analysis');
 const byName=new Map(sheets.map(s=>[s.name,s]));
 const usedIds=new Set<string>(),usedSheetIds=new Set<string>();
 for(const node of declared){
  const name=attribute(node,'name'),id=attribute(node,'sheetId'),rid=attribute(node,'id',OFFICE),sheet=byName.get(name??'');
  if(!name||!sheet||!id||!/^\d+$/.test(id)||Number(id)<1||usedSheetIds.has(id)||!rid||usedIds.has(rid)||(attribute(node,'state')!==undefined&&attribute(node,'state')!=='visible'))refuse('Invalid worksheet identity');
  const ownedSheet=sheet!;
  const relationshipId=rid!,sheetId=id!;
  const edge=rels.find(r=>r.id===relationshipId);
  if(!edge||edge.external||edge.type!==SHEET_REL||edge.resolved!==ownedSheet.part)refuse('Unresolved worksheet relationship');
  usedIds.add(relationshipId);usedSheetIds.add(sheetId);
 }
 if(rels.filter(r=>r.type===SHEET_REL).length!==sheets.length||names.filter(n=>/^xl\/worksheets\/[^/]+\.xml$/.test(n)).length!==sheets.length)refuse('Unlisted worksheet part');
 if(elements(workbook.root,'externalReferences',S).length||elements(workbook.root,'definedName',S).length||names.some(n=>/^xl\/(externalLinks|tables|pivotTables|connections|metadata|model|queryTables|charts|pivotCache)\//.test(n)))refuse('Unsupported formula-bearing package part');
 const chainDoc=parseXml(pkg.text(part)),root=chainDoc.root;
 if(root.localName!=='calcChain'||root.namespaceURI!==S||Object.keys(root.attributes).some(k=>(k!=='xmlns'&&!k.startsWith('xmlns:'))||root.attributeNamespaces[k]!=='http://www.w3.org/2000/xmlns/')||root.text.trim()||!root.children.length)refuse('Invalid calculation-chain root');
 const chainCells=new Set<string>();
 for(const child of root.children){
  const r=attribute(child,'r'),id=attribute(child,'i');
  if(child.localName!=='c'||child.namespaceURI!==S||child.children.length||child.text.trim()||!exact(child,['r','i','l','s','a','t'])||!r||!cellReference(r)||!id||!usedSheetIds.has(id)||!['l','s','a','t'].every(k=>{const v=attribute(child,k);return v===undefined||['0','1','true','false'].includes(v)}))refuse('Unsupported calculation-chain entry');
  const key=`${id}!${r}`;if(chainCells.has(key))refuse('Duplicate calculation-chain entry');chainCells.add(key);
 }
 const graph=new Map<string,Set<string>>(),existing=new Set<string>();
 for(const sheet of sheets){
  if(sheet.document.root.localName!=='worksheet'||sheet.document.root.namespaceURI!==S||elements(sheet.document.root,'tableParts',S).length||elements(sheet.document.root,'extLst',S).length)refuse('Unsupported worksheet structure');
  const cells=elements(sheet.document.root,'c',S);
  if(cells.length!==sheet.refs.size||elements(sheet.document.root,'f',S).length!==cells.reduce((n,c)=>n+children(c,'f',S).length,0))refuse('Incomplete worksheet cell or formula index');
  for(const cell of cells){
   const ref=attribute(cell,'r');if(!ref||!sheet.refs.has(ref)||!cellReference(ref))refuse('Invalid worksheet cell reference');
   const key=`${sheet.name}!${ref}`;if(existing.has(key))refuse('Duplicate worksheet cell');existing.add(key);
   if(existing.size>100000)refuse('Too many cells for dependency analysis');
   const formulas=children(cell,'f',S);if(formulas.length>1)refuse('Duplicate formula');
   if(!formulas.length)continue;
   const f=formulas[0]!,expression=f.text;
   if(Object.keys(f.attributes).length||f.children.length||!expression||!/^[\x20-\x7e]+$/.test(expression)||!/^=?[A-Za-z0-9_!.$()+\-*/\s]+$/.test(expression)||/[A-Za-z_]\w*\s*\(/.test(expression))refuse('Unsupported formula topology');
   let refs:ReturnType<typeof analyzeFormulaReferences>=[];
   try{refs=analyzeFormulaReferences(expression);}catch{refuse('Unresolved or unsupported formula reference');}
   const dependencies=new Set<string>();
   for(const r of refs){
    if(r.axis!=='cell'||r.first.row!==r.last.row||r.first.column!==r.last.column||!byName.has(r.sheet||sheet.name))refuse('Unsupported formula reference');
    // The analyser's spans are UTF-8 byte offsets. The column/row values are
    // already parsed and range-checked; reconstruct a cell key without slicing
    // the UTF-16 source string at byte offsets.
    const col=(n:number)=>{let value='';while(n){n--;value=String.fromCharCode(65+n%26)+value;n=Math.floor(n/26);}return value;};
    const ref=`${col(r.first.column)}${r.first.row}`;
    dependencies.add(`${r.sheet||sheet.name}!${ref}`);
   }
   graph.set(key,dependencies);
   if(graph.size>10000)refuse('Too many formulas for dependency analysis');
  }
 }
 if(!existing.has(`${editedSheet}!${editedRef}`)||[...graph.values()].some(refs=>[...refs].some(ref=>!existing.has(ref))))refuse('Incomplete formula graph');
 for(const child of root.children){
  const id=attribute(child,'i')!,ref=attribute(child,'r')!,sheet=declared.find(s=>attribute(s,'sheetId')===id)!;
  if(!graph.has(`${attribute(sheet,'name')}!${ref}`))refuse('Calculation-chain entry has no formula');
 }
 const visiting=new Set<string>(),depths=new Map<string,number>();
 const visit=(key:string,level:number):number=>{
  if(level>128)refuse('Formula dependency depth exceeds limit');
  if(visiting.has(key))refuse('Circular formula dependency');
  const known=depths.get(key);if(known!==undefined)return known;
  visiting.add(key);
  let depth=0;
  for(const ref of graph.get(key)??[])if(graph.has(ref))depth=Math.max(depth,visit(ref,level+1)+1);
  visiting.delete(key);
  if(depth>128)refuse('Formula dependency depth exceeds limit');
  depths.set(key,depth);
  return depth;
 };
 for(const key of graph.keys())visit(key,0);
 const affected=new Set([`${editedSheet}!${editedRef}`]);
 for(let changed=true;changed;){changed=false;for(const [key,refs] of graph)if(!affected.has(key)&&[...refs].some(ref=>affected.has(ref))){affected.add(key);changed=true;}}
 affected.delete(`${editedSheet}!${editedRef}`);
 return {part,relationshipId:chain.id,affected};
}
