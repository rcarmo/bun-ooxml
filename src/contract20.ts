import {posix} from 'node:path';
import {Workbook,type Cell} from './xlsx/index.ts';
import {Presentation,type Slide,type TextAnchor,type InspectedParagraph,type TableGeometry} from './pptx/index.ts';
import {OoxmlError} from './errors.ts';
import {relationshipPath,sameBytes,type OpcPackage} from './opc/package.ts';
import {parseXml,elements,attribute,applyEdits,escapeText,type XmlElement} from './xml/index.ts';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',CT='http://schemas.openxmlformats.org/package/2006/content-types';
type Input=string|Uint8Array|ArrayBuffer;
const TABLE_URI='http://schemas.openxmlformats.org/drawingml/2006/table';
function direct(n:XmlElement,local:string,ns:string):XmlElement|undefined{return n.children.find(c=>c.localName===local&&c.namespaceURI===ns);}
function tablePayload(frame:XmlElement):XmlElement{const graphic=direct(frame,'graphic',A),data=graphic&&direct(graphic,'graphicData',A);if(!data||data.attributes.uri!==TABLE_URI||data.children.length!==1||data.children[0]!.localName!=='tbl'||data.children[0]!.namespaceURI!==A)fail('PPTX_TABLE_STRUCTURE_UNSUPPORTED','Foreign or non-direct table payload');return data.children[0]!;}
function fail(code:string,message:string):never{throw new OoxmlError(code,message);}
function validText(text:string):void {if(typeof text!=='string')fail('PPTX_ARGUMENT_INVALID','Text must be a Unicode string');try{escapeText(text);}catch(e){if(e instanceof OoxmlError&&e.code==='XML_INVALID_CHAR')fail('PPTX_ARGUMENT_INVALID','Text contains an invalid XML character');throw e;}}
function captured(pkg:OpcPackage):Map<string,Uint8Array>{return new Map(pkg.names().map(p=>[p,pkg.get(p)!]));}
function unchanged(pkg:OpcPackage,before:Map<string,Uint8Array>):boolean{return pkg.names().length===before.size&&[...before].every(([p,b])=>{const now=pkg.get(p);return !!now&&sameBytes(now,b);});}
const opaque=[{id:'cache1',type:'urn:contract20:opaque/chart',part:'xl/charts/cache-boundary.xml',contentType:'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'},{id:'cache2',type:'urn:contract20:opaque/external',part:'xl/externalLinks/cache-boundary.xml',contentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'}];
function cacheAdmission(pkg:OpcPackage,optIn:boolean):void {
 const root=pkg.relationships(),main=pkg.mainPart(),overrides=elements(parseXml(pkg.text('[Content_Types].xml')),'Override',CT);
 const rootDocument=root.filter(r=>r.type===R+'/officeDocument'&&!r.external&&r.resolved===main);
 if(rootDocument.length!==1)fail('xlsx-cache-topology-unsupported','Exactly one root workbook edge is required');
 const workbookRels=pkg.relationships(main),sheetNodes=elements(parseXml(pkg.text(main)),'sheet',S),sheetIds=sheetNodes.map(n=>attribute(n,'id',R));
 if(sheetIds.some(id=>!id)||new Set(sheetIds).size!==sheetIds.length)fail('xlsx-cache-topology-unsupported','Invalid worksheet relationship identity');
 const sheetParts=new Set<string>();for(const id of sheetIds){const rel=workbookRels.find(r=>r.id===id);if(!rel||rel.external||rel.type!==R+'/worksheet'||!rel.resolved||sheetParts.has(rel.resolved))fail('xlsx-cache-topology-unsupported','Ambiguous worksheet edge');sheetParts.add(rel.resolved);}
 const defaultTypes=elements(parseXml(pkg.text('[Content_Types].xml')),'Default',CT),partType=(part:string)=>overrides.find(n=>n.attributes.PartName==='/'+part)?.attributes.ContentType??defaultTypes.find(n=>n.attributes.Extension===part.split('.').at(-1))?.attributes.ContentType;
 for(const role of ['styles','sharedStrings'])if(workbookRels.filter(r=>r.type===R+'/'+role).length>1)fail('xlsx-cache-topology-unsupported','Extra workbook support edge');
 const present=opaque.some(o=>pkg.get(o.part)||root.some(r=>r.id===o.id||r.type===o.type));
 if(root.length!==(present?3:1))fail('xlsx-cache-topology-unsupported','Extra root semantic edge is outside the exact profile');
 if(present&&!optIn)fail('xlsx-cache-topology-unsupported','Opaque records require the exact opt-in retention profile');
 if(present)for(const o of opaque){const rel=root.filter(r=>r.id===o.id);if(rel.length!==1||rel[0]!.external||rel[0]!.type!==o.type||rel[0]!.target!==o.part||rel[0]!.resolved!==o.part||!pkg.get(o.part)||pkg.get(relationshipPath(o.part)))fail('xlsx-cache-topology-unsupported','Opaque record edge or outgoing custody mismatch');const ct=overrides.filter(n=>n.attributes.PartName==='/'+o.part);if(ct.length!==1||ct[0]!.attributes.ContentType!==o.contentType)fail('xlsx-cache-topology-unsupported','Opaque record content type mismatch');}
 for(const p of pkg.names().filter(n=>n.endsWith('.rels'))){const owner=p==='_rels/.rels'?'':posix.join(posix.dirname(posix.dirname(p)),posix.basename(p).slice(0,-5));for(const rel of pkg.relationships(owner)){
  if(owner===''){
   if(rel.type===R+'/officeDocument'&&!rel.external&&rel.resolved===main)continue;
   if(present&&opaque.some(o=>rel.id===o.id&&rel.type===o.type&&rel.target===o.part&&rel.resolved===o.part&&!rel.external))continue;
   fail('xlsx-cache-topology-unsupported','Unexpected root relationship');
  }
  if(owner!==main||rel.external||![R+'/worksheet',R+'/sharedStrings',R+'/styles'].includes(rel.type)||(rel.resolved&&opaque.some(o=>o.part===rel.resolved)))fail('xlsx-cache-topology-unsupported','Unknown owner or semantic cache dependency');
  const role=rel.type.slice(R.length+1),target=rel.resolved!;
  if(role==='worksheet'&&(!sheetIds.includes(rel.id)||!sheetParts.has(target)||partType(target)!=='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'||parseXml(pkg.text(target)).root.localName!=='worksheet'))fail('xlsx-cache-topology-unsupported','Worksheet owner/target role mismatch');
  if(role==='styles'||role==='sharedStrings'){const expected=role==='styles'?'styleSheet':'sst';if(partType(target)!=='application/vnd.openxmlformats-officedocument.spreadsheetml.'+role+'+xml'||parseXml(pkg.text(target)).root.localName!==expected||parseXml(pkg.text(target)).root.namespaceURI!==S)fail('xlsx-cache-topology-unsupported','Workbook support target role mismatch');}
 }}
 for(const p of pkg.names().filter(n=>/^xl\/(?:charts|externalLinks)\//.test(n)))if(!opaque.some(o=>o.part===p))fail('xlsx-cache-topology-unsupported','Unknown opaque part');
}
/** Opt-in retained scalar edits, all-ordinary caches and exact inert retention. */
export class ContractWorkbook {
 private snapshot:Map<string,Uint8Array>;
 private constructor(public workbook:Workbook){this.snapshot=captured(workbook.package);}
 static async open(input:Input):Promise<ContractWorkbook>{return new ContractWorkbook(await Workbook.open(input));}
 get package():OpcPackage{return this.workbook.package;}
 get sheetnames():string[]{return this.workbook.sheetnames;}
 readCell(sheet:string,reference:string):Cell|undefined{return this.workbook.worksheet(sheet).getCell(reference);}
 setCellValue(sheet:string,reference:string,value:string|number|boolean,options:{opaqueRetention?:boolean}={}):void {
  if(!unchanged(this.package,this.snapshot))fail('xlsx-stale-workbook','Workbook changed outside this profile');
  if(typeof value!=='string'&&typeof value!=='number'&&typeof value!=='boolean')fail('xlsx-value-invalid','Expected a scalar value');
  if(typeof value==='number'&&!Number.isFinite(value))fail('xlsx-value-invalid','Expected a finite number');
  if(typeof value==='string')escapeText(value);
  if(typeof options.opaqueRetention!=='undefined'&&typeof options.opaqueRetention!=='boolean')fail('xlsx-cache-topology-unsupported','Invalid retention option');
  const cell=this.readCell(sheet,reference);
  // A no-op is only an existing ordinary cell of the same semantic scalar type.
  if(cell&&cell.kind!=='formula'&&cell.kind!=='blank'&&cell.value===value)return;
  if(cell?.kind==='formula'){
   const main=this.package.mainPart(),edge=this.package.relationships(main).find(r=>r.type===R+'/worksheet'&&r.resolved&&elements(parseXml(this.package.text(main)),'sheet',S).some(n=>n.attributes.name===sheet&&attribute(n,'id',R)===r.id));
   const selected=edge?.resolved?elements(parseXml(this.package.text(edge.resolved)),'c',S).find(c=>c.attributes.r===cell.ref):undefined,type=selected?.children.find(n=>n.localName==='f'&&n.namespaceURI===S)?.attributes.t;
   if(type==='shared')fail('xlsx-shared-formula-edit-unsupported','Direct shared formula overwrite is unsupported');
   if(type==='array')fail('xlsx-array-formula-edit-unsupported','Direct array formula overwrite is unsupported');
   fail('xlsx-formula-edit-unsupported','Direct formula overwrite is unsupported');
  }
  cacheAdmission(this.package,options.opaqueRetention===true);
  const root=parseXml(this.package.text(this.package.mainPart())),rels=this.package.relationships(this.package.mainPart());
  const worksheets=elements(root,'sheet',S).map(n=>{const id=attribute(n,'id',R),r=rels.find(r=>r.id===id);if(!r?.resolved)fail('xlsx-workbook-invalid','Worksheet relationship missing');return r.resolved;});
  for(const part of worksheets)for(const formula of elements(parseXml(this.package.text(part)),'f',S))if(formula.attributes.t!==undefined&&formula.attributes.t!=='normal'||['si','ref','dt2D','dtr','r1','r2'].some(k=>Object.hasOwn(formula.attributes,k)))fail('xlsx-cache-topology-unsupported','Changed-input sweep requires ordinary formula topology');
  try{this.package.transaction(()=>{
   this.workbook.worksheet(sheet).setCellValue(reference,value);
   // The legacy editor empties caches; this profile removes the cache nodes.
   for(const part of worksheets){const xml=this.package.text(part),doc=parseXml(xml),edits=elements(doc,'c',S).flatMap(c=>c.children.some(n=>n.localName==='f'&&n.namespaceURI===S)?c.children.filter(n=>n.localName==='v'&&n.namespaceURI===S).map(n=>({start:n.start,end:n.end,value:''})):[]);if(edits.length)this.package.set(part,applyEdits(xml,edits));}
   this.package.toBytes();
  });}catch(error){this.workbook=Workbook.fromPackage(this.package);throw error;}
  this.workbook=Workbook.fromPackage(this.package);
  this.snapshot=captured(this.package);
 }
 toBytes():Uint8Array{return this.package.toBytes();}
 save(path:string):Promise<void>{return this.package.save(path);}
}

export interface ContractTextAnchor {readonly kind:'contract20-text';readonly part:string;readonly text:string;}
export interface ContractParagraph {text:string;runs:InspectedParagraph['runs'];anchor:ContractTextAnchor;}
type TextCustody={slide:Slide;anchor:TextAnchor;partBytes:Uint8Array;relationships?:Uint8Array};
function retainedPart(pkg:OpcPackage,part:string,before:Uint8Array):boolean{const now=pkg.get(part);return !!now&&sameBytes(now,before);}
export interface ContractTableCell {readonly part:string;readonly shapeId:number;readonly row:number;readonly column:number;}
type CellCustody={slide:Slide;partBytes:Uint8Array;relationships?:Uint8Array;generation:number;cell:XmlElement;table:XmlElement};
function retainedRelationships(pkg:OpcPackage,part:string,before?:Uint8Array):boolean{const now=pkg.get(relationshipPath(part));return before?!!now&&sameBytes(now,before):now===undefined;}
/** Opt-in part-scoped issued targets and table edit-preflight separation. */
export class ContractPresentation {
 private readonly anchors=new WeakMap<ContractTextAnchor,TextCustody>();
 private readonly cells=new WeakMap<ContractTableCell,CellCustody>();
 private constructor(readonly presentation:Presentation){}
 static create():ContractPresentation{return new ContractPresentation(Presentation.create());}
 static async open(input:Input):Promise<ContractPresentation>{return new ContractPresentation(await Presentation.open(input));}
 get package():OpcPackage{return this.presentation.package;}
 get slides():Slide[]{return this.presentation.slides;}
 addTextSlide(title:string,subtitle?:string):Slide{validText(title);if(subtitle!==undefined)validText(subtitle);return this.presentation.addTextSlide(title,subtitle);}
 inspectText(slideIndex:number):ContractParagraph[]{const slide=this.slides[slideIndex];if(!slide)fail('PPTX_ARGUMENT_INVALID','Unknown slide');return slide.inspectText('contract20').map(p=>{const anchor:ContractTextAnchor=Object.freeze({kind:'contract20-text',part:slide.partName,text:p.text});this.anchors.set(anchor,{slide,anchor:p.anchor,partBytes:this.package.get(slide.partName)!,relationships:this.package.get(relationshipPath(slide.partName))});return {text:p.text,runs:p.runs,anchor};});}
 replaceTextAt(anchor:ContractTextAnchor,find:string,replace:string):void{const owned=this.anchors.get(anchor);if(!owned||owned.anchor.version!==this.presentation.currentSlideVersion(owned.slide.partName)||!retainedRelationships(this.package,owned.slide.partName,owned.relationships)||!retainedPart(this.package,owned.slide.partName,owned.partBytes))fail('PPTX_STALE_ANCHOR','Target was not issued by this current owning part');validText(find);validText(replace);owned.slide.replaceTextAt(owned.anchor,find,replace);}
 readNotes(slideIndex:number):{hasNotes:boolean;text:string}{const slide=this.slides[slideIndex];if(!slide)fail('PPTX_ARGUMENT_INVALID','Unknown slide');try{return {hasNotes:true,text:slide.readNotesText()};}catch(e){if(e instanceof OoxmlError&&e.code==='PPTX_NOTES_MISSING')return {hasNotes:false,text:''};throw e;}}
 addTable(slideIndex:number,rows:number,columns:number,geometry:TableGeometry):void{const s=this.slides[slideIndex];if(!s)fail('PPTX_ARGUMENT_INVALID','Unknown slide');s.addTable(rows,columns,geometry);}
 selectTableCell(slideIndex:number,shapeId:number,row:number,column:number):ContractTableCell {
  const slide=this.slides[slideIndex];if(!slide||!Number.isSafeInteger(shapeId)||!Number.isSafeInteger(row)||!Number.isSafeInteger(column)||row<0||column<0)fail('PPTX_ARGUMENT_INVALID','Invalid table target');
  const bytes=this.package.get(slide.partName)!,doc=parseXml(new TextDecoder('utf-8',{fatal:true}).decode(bytes)),tree=direct(direct(doc.root,'cSld',P)??doc.root,'spTree',P),frames=tree?.children.filter(f=>f.localName==='graphicFrame'&&f.namespaceURI===P&&direct(direct(f,'nvGraphicFramePr',P)??f,'cNvPr',P)?.attributes.id===String(shapeId))??[];
  if(frames.length!==1)fail('PPTX_ARGUMENT_INVALID','Table frame must be unique');const table=tablePayload(frames[0]!),rows=table.children.filter(n=>n.localName==='tr'&&n.namespaceURI===A),cell=rows[row]?.children.filter(n=>n.localName==='tc'&&n.namespaceURI===A)[column];if(!cell)fail('PPTX_ARGUMENT_INVALID','Table cell missing');
  const target:ContractTableCell=Object.freeze({part:slide.partName,shapeId,row,column});this.cells.set(target,{slide,partBytes:bytes,relationships:this.package.get(relationshipPath(slide.partName)),generation:this.presentation.currentSlideVersion(slide.partName),cell,table});return target;
 }
 setTableCellText(target:ContractTableCell,text:string):void {
  const owned=this.cells.get(target);if(!owned||owned.generation!==this.presentation.currentSlideVersion(owned.slide.partName)||!retainedRelationships(this.package,owned.slide.partName,owned.relationships)||!retainedPart(this.package,owned.slide.partName,owned.partBytes))fail('PPTX_STALE_TABLE_HANDLE','Target was not issued by this current owning part');validText(text);
  const {table,cell}=owned,grid=table.children.find(n=>n.localName==='tblGrid'&&n.namespaceURI===A)?.children.filter(n=>n.localName==='gridCol'&&n.namespaceURI===A)??[],rows=table.children.filter(n=>n.localName==='tr'&&n.namespaceURI===A);let merged=false;
  if(!grid.length||!rows.length)fail('PPTX_TABLE_STRUCTURE_UNSUPPORTED','Empty table');
  for(const row of rows){let occupancy=0;for(const c of row.children.filter(n=>n.localName==='tc'&&n.namespaceURI===A)){let width=1;for(const key of ['gridSpan','rowSpan']){const raw=c.attributes[key];if(raw!==undefined){if(!/^[1-9][0-9]*$/.test(raw)||!Number.isSafeInteger(Number(raw)))fail('PPTX_TABLE_STRUCTURE_UNSUPPORTED','Invalid span');if(Number(raw)>1)merged=true;if(key==='gridSpan')width=Number(raw);}}for(const key of ['hMerge','vMerge']){const raw=c.attributes[key];if(raw!==undefined){if(!['0','1','true','false'].includes(raw))fail('PPTX_TABLE_STRUCTURE_UNSUPPORTED','Invalid merge flag');if(raw==='1'||raw==='true')merged=true;}}occupancy+=width;}if(occupancy!==grid.length)fail('PPTX_TABLE_STRUCTURE_UNSUPPORTED','Logical row occupancy mismatch');}
  if(merged)fail('PPTX_TABLE_MERGE_UNSUPPORTED','Merged table text editing is not supported');
  const leaves=elements(cell,'t',A),paragraphs=elements(cell,'p',A);let next:string;const xml=this.package.text(owned.slide.partName);
  const frame=elements(parseXml(xml),'graphicFrame',P).find(f=>direct(direct(f,'nvGraphicFramePr',P)??f,'cNvPr',P)?.attributes.id===String(target.shapeId));if(!frame)fail('PPTX_STALE_TABLE_HANDLE','Table frame disappeared');tablePayload(frame);
  if(paragraphs.length!==1||leaves.length>1||elements(cell,'br',A).length||elements(cell,'fld',A).length)fail('PPTX_UNSUPPORTED_TEXT_TOPOLOGY','Retained cell text requires a single plain text leaf or empty paragraph');
  if(leaves.length===1&&(!direct(cell,'txBody',A)||!paragraphs[0]!.children.some(n=>n.localName==='r'&&n.namespaceURI===A&&direct(n,'t',A)===leaves[0])))fail('PPTX_UNSUPPORTED_TEXT_TOPOLOGY','Text leaf is not an owned direct run');
  if(leaves.length===1&&!leaves[0]!.selfClosing&&paragraphs.length===1&&!elements(cell,'br',A).length&&!elements(cell,'fld',A).length){const leaf=leaves[0]!;next=applyEdits(xml,[{start:leaf.openEnd,end:leaf.closeStart,value:escapeText(text)}]);}
  else {const frames=elements(parseXml(xml),'graphicFrame',P).filter(f=>elements(f,'tbl',A).length===1),index=frames.findIndex(f=>elements(f,'cNvPr',P).some(n=>n.attributes.id===String(target.shapeId))),table=owned.slide.tables[index];if(!table)fail('PPTX_STALE_TABLE_HANDLE','Table identity changed');table.cell(target.row,target.column).text=text;return;}
  if(next===xml)return;this.package.transaction(()=>{this.package.set(owned.slide.partName,next);this.package.toBytes();});this.presentation.bumpSlideVersion(owned.slide.partName);
 }
 save(path:string):Promise<void>{return this.package.save(path);}
}
