import {OoxmlError} from '../errors.ts';
import {OpcPackage,type Relationship,relationshipPath} from '../opc/package.ts';
import {getContentType} from '../opc/content-types.ts';
import {parseXml,attribute,type XmlElement} from '../xml/index.ts';
const S='http://schemas.openxmlformats.org/spreadsheetml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const XML='http://www.w3.org/XML/1998/namespace',XMLNS='http://www.w3.org/2000/xmlns/';
const SHEET='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',COMMENTS='application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',VML='application/vnd.openxmlformats-officedocument.vmlDrawing';
export interface WorksheetComment {reference:string;authorId:number;author:string;text:string}
export interface CommentPartReference {id:string;type:string;part:string}
export interface WorksheetComments {worksheetPart:string;commentsRelationship:CommentPartReference|null;vmlRelationship:CommentPartReference|null;comments:WorksheetComment[]}
function fail(message:string):never{throw new OoxmlError('xlsx-comments-unsupported',message);}
const is=(n:XmlElement,name:string)=>n.localName===name&&n.namespaceURI===S;
function attrs(n:XmlElement,allowed:string[]){for(const key of Object.keys(n.attributes)){if(n.attributeNamespaces[key]===XMLNS)continue;if(n.attributeNamespaces[key]!==''||!allowed.includes(key))fail('Unsupported comment metadata attribute');}}
function gaps(n:XmlElement,xml:string){if(n.selfClosing)return;let at=n.openEnd;for(const c of n.children){if(!/^[ \t\r\n]*$/.test(xml.slice(at,c.start)))fail('Unsupported mixed or lexical comment metadata');at=c.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(at,n.closeStart)))fail('Unsupported mixed or lexical comment metadata');}
function leaf(n:XmlElement,xml:string):string{if(n.children.length||!n.selfClosing&&xml.slice(n.openEnd,n.closeStart).includes('<'))fail('Nested or lexical text is unsupported');return n.text;}
function textLeaf(n:XmlElement,xml:string):string{
 if(!is(n,'t'))fail('Expected SpreadsheetML text');
 for(const key of Object.keys(n.attributes))if(n.attributeNamespaces[key]!==XMLNS&&(n.attributeNamespaces[key]!==XML||key.split(':').at(-1)!=='space'))fail('Unsupported text attribute');
 const space=attribute(n,'space',XML);if(space!==undefined&&space!=='preserve'&&space!=='default')fail('Invalid text whitespace metadata');return leaf(n,xml);
}
function readText(n:XmlElement,xml:string):string{
 attrs(n,[]);gaps(n,xml);
 if(n.children.length===1&&is(n.children[0]!,'t'))return textLeaf(n.children[0]!,xml);
 if(!n.children.length)return '';
 if(n.children.some(c=>!is(c,'r')))fail('Unsupported or mixed rich comment text');
 return n.children.map(run=>{
  attrs(run,[]);gaps(run,xml);const children=run.children;
  if(children.length===2&&is(children[0]!,'rPr')){
   const pr=children[0]!;attrs(pr,[]);gaps(pr,xml);const seen=new Set<string>();
   for(const value of pr.children){if(value.namespaceURI!==S||!['rFont','charset','family','b','i','strike','outline','shadow','condense','extend','color','sz','u','vertAlign','scheme'].includes(value.localName)||seen.has(value.localName))fail('Unsupported run properties');seen.add(value.localName);attrs(value,value.localName==='color'?['auto','indexed','rgb','theme','tint']:['val']);gaps(value,xml);if(value.children.length)fail('Nested run property');}
  }else if(children.length!==1)fail('Expected one text leaf and optional first run properties');
  return textLeaf(children.at(-1)!,xml);
 }).join('');
}
function reference(rel:Relationship,pkg:OpcPackage,type:string,mime:string):CommentPartReference{
 if(rel.type!==type||rel.external||!rel.resolved||rel.target.includes('#')||rel.target.includes('?')||!pkg.get(rel.resolved)||getContentType(pkg,rel.resolved)!==mime)fail('Unsupported or missing comment/VML dependency');
 return {id:rel.id,type:rel.type,part:rel.resolved};
}
/** Read current Transitional worksheet comment metadata without interpreting VML shapes or mutating the package. */
export function inspectWorksheetComments(pkg:OpcPackage,worksheetPart:string):WorksheetComments{
 if(!(pkg instanceof OpcPackage)||typeof worksheetPart!=='string'||!worksheetPart||/[\\:\u0000-\u001f]/.test(worksheetPart)||worksheetPart.split('/').some(p=>!p||p==='.'||p==='..'))fail('Expected a package and canonical worksheet part name');
 if(!pkg.get(worksheetPart)||getContentType(pkg,worksheetPart)!==SHEET)fail('Expected a worksheet content type');
 const sheet=parseXml(pkg.text(worksheetPart));if(!is(sheet.root,'worksheet'))fail('Expected Transitional worksheet root');
 const markers=sheet.elements.filter(n=>n.localName==='legacyDrawing');if(markers.length>1||markers.some(n=>n.parent!==sheet.root||n.namespaceURI!==S))fail('Ambiguous legacy drawing metadata');
 let legacyId:string|undefined;
 if(markers.length){const marker=markers[0]!;for(const key of Object.keys(marker.attributes))if(marker.attributeNamespaces[key]!==XMLNS&&(marker.attributeNamespaces[key]!==R||key.split(':').at(-1)!=='id'))fail('Invalid legacy drawing relationship attribute');if(marker.children.length||!marker.selfClosing&&!/^[ \t\r\n]*$/.test(pkg.text(worksheetPart).slice(marker.openEnd,marker.closeStart)))fail('Invalid legacy drawing content');legacyId=attribute(marker,'id',R);if(!legacyId)fail('Missing legacy drawing relationship ID');}
 const relationships=pkg.relationships(worksheetPart);
 // Selected relationship attributes must be unqualified, not merely local-name lookalikes.
 const relPath=relationshipPath(worksheetPart);if(pkg.get(relPath)){const rels=parseXml(pkg.text(relPath));for(const n of rels.root.children)attrs(n,['Id','Type','Target','TargetMode']);}
 const vmls=relationships.filter(r=>r.type===R+'/vmlDrawing'),commentRels=relationships.filter(r=>r.type===R+'/comments');
 if(commentRels.length>1||vmls.length>1||Boolean(legacyId)!==Boolean(vmls.length)||legacyId&&vmls[0]!.id!==legacyId)fail('Ambiguous comment or VML relationships');
 const vmlRelationship=vmls.length?reference(vmls[0]!,pkg,R+'/vmlDrawing',VML):null;
 const commentsRelationship=commentRels.length?reference(commentRels[0]!,pkg,R+'/comments',COMMENTS):null;
 if(commentsRelationship&&vmlRelationship&&(commentsRelationship.id===vmlRelationship.id||commentsRelationship.part===vmlRelationship.part))fail('Comments and VML must be separate dependencies');
 const result:WorksheetComments={worksheetPart,commentsRelationship,vmlRelationship,comments:[]};if(!commentsRelationship)return result;
 const xml=pkg.text(commentsRelationship.part),doc=parseXml(xml),root=doc.root;if(!is(root,'comments'))fail('Expected Transitional comments root');attrs(root,[]);gaps(root,xml);
 if(root.children.length!==2||!is(root.children[0]!,'authors')||!is(root.children[1]!,'commentList'))fail('Expected authors then commentList only');
 const authorsNode=root.children[0]!,list=root.children[1]!;attrs(authorsNode,[]);gaps(authorsNode,xml);attrs(list,[]);gaps(list,xml);
 if(authorsNode.children.length>10000||list.children.length>10000)fail('Comment or author count exceeds 10000');
 const authors=authorsNode.children.map(n=>{if(!is(n,'author'))fail('Expected author text');attrs(n,[]);return leaf(n,xml);});
 const seen=new Set<string>();
 for(const n of list.children){if(!is(n,'comment'))fail('Expected comment');attrs(n,['ref','authorId','shapeId']);gaps(n,xml);const ref=attribute(n,'ref'),id=attribute(n,'authorId'),match=/^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(ref??'');if(!match)fail('Expected canonical A1 cell reference');let col=0;for(const c of match[1]!)col=col*26+c.charCodeAt(0)-64;if(col>16384||Number(match[2])>1048576||seen.has(ref!))fail('Duplicate or out-of-range comment reference');seen.add(ref!);
  if(!/^(0|[1-9][0-9]*)$/.test(id??'')||!Number.isSafeInteger(Number(id))||Number(id)>=authors.length)fail('Invalid comment author index');
  if(n.children.length!==1||!is(n.children[0]!,'text'))fail('Expected exactly one comment text');
  result.comments.push({reference:ref!,authorId:Number(id),author:authors[Number(id)]!,text:readText(n.children[0]!,xml)});
 }
 return result;
}
