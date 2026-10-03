import {OoxmlError} from '../errors.ts';
import {type OpcPackage,relationshipPath,type Relationship} from '../opc/package.ts';
import {removeRelationship,removePart} from '../opc/graph.ts';
import {getContentType} from '../opc/content-types.ts';
import {inspectPictures} from './pictures.ts';
import {parseXml,elements,attribute,applyEdits,type XmlElement} from '../xml/index.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',IMAGE=R+'/image';
export type PictureDeleteOptions={collectMedia?:boolean};
export type PictureDeleteReceipt={shapeId:number;partName:string;removedRelationships:string[];removedMedia:string[]};
function fail(message:string):never {throw new OoxmlError('PPTX_PICTURE_UNSUPPORTED',message);}
function request(id:number,options:PictureDeleteOptions):boolean{
 if(!Number.isInteger(id)||id<1||id>2147483647)fail('Deletion requires a bounded picture ID');
 if(!options||typeof options!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(options)))fail('Deletion options require plain data');
 let collect=false;for(const k of Reflect.ownKeys(options)){if(k!=='collectMedia')fail('Unknown deletion option');const d=Object.getOwnPropertyDescriptor(options,k)!;if(!('value'in d)||typeof d.value!=='boolean')fail('Collection requires a Boolean data property');collect=d.value;}return collect;
}
function walk(node:XmlElement):XmlElement[]{return [node,...node.children.flatMap(walk)];}
function incoming(pkg:OpcPackage,target:string):boolean{
 const owners=['',...pkg.names().filter(n=>n!=='[Content_Types].xml'&&!n.endsWith('.rels'))];
 return owners.some(owner=>pkg.relationships(owner).some(r=>!r.external&&r.resolved===target));
}
/** Delete one literal picture span; collect only bounded, proven-unused dependencies. */
export function deletePicture(pkg:OpcPackage,part:string,id:number,options:PictureDeleteOptions={}):PictureDeleteReceipt{
 const collect=request(id,options),info=inspectPictures(pkg,part).find(p=>p.shapeId===id);if(!info)throw new OoxmlError('PPTX_PICTURE_NOT_FOUND','No picture with the exact ID');
 const source=pkg.text(part),doc=parseXml(source),picture=elements(doc,'pic',P).find(n=>n.children.find(c=>c.namespaceURI===P&&c.localName==='nvPicPr')?.children.some(c=>c.namespaceURI===P&&c.localName==='cNvPr'&&Number(attribute(c,'id'))===id))!;
 for(const n of [...elements(doc,'stCxn','http://schemas.openxmlformats.org/drawingml/2006/main'),...elements(doc,'endCxn','http://schemas.openxmlformats.org/drawingml/2006/main')])if(Number(attribute(n,'id'))===id)fail('Picture is an attached connector endpoint');
 const relationships=pkg.relationships(part),dependencies:Relationship[]=[];
 for(const n of walk(picture))for(const local of ['embed','link','id']){
  const value=attribute(n,local,R);if(value===undefined)continue;const matches=relationships.filter(r=>r.id===value);
  if(matches.length!==1||local!=='id'&&matches[0]!.type!==IMAGE)fail('Ambiguous picture dependency');
  if(matches[0]!.type===IMAGE&&!dependencies.some(r=>r.id===value))dependencies.push(matches[0]!);
 }
 const next=applyEdits(source,[{start:picture.start,end:picture.end,value:''}]),remaining=parseXml(next);
 // Unknown text/comment grammars cannot prove reference absence. XML declaration is harmless.
 const lexical=next.replace(/^\uFEFF?\s*<\?xml\s[^?]*\?>/,'');
 const canCollect=collect&&!/<!--|<\?|<!\[CDATA\[/.test(lexical)&&remaining.elements.every(n=>!n.directText.trim()||n.namespaceURI==='http://schemas.openxmlformats.org/drawingml/2006/main'&&n.localName==='t');
 return pkg.transaction(()=>{
  pkg.set(part,next);const removedRelationships:string[]=[],removedMedia:string[]=[];
  if(canCollect){
   for(const r of dependencies){if(remaining.elements.some(n=>Object.values(n.attributes).includes(r.id)))continue;removeRelationship(pkg,part,r.id);removedRelationships.push(r.id);}
   for(const r of dependencies){
    if(r.external||!r.resolved||removedMedia.includes(r.resolved)||!removedRelationships.includes(r.id))continue;
    const target=r.resolved,type=getContentType(pkg,target);
    if(!/^ppt\/media\/[^/]+\.(?:png|jpe?g|svg)$/i.test(target)||!type?.startsWith('image/')||pkg.get(relationshipPath(target))||incoming(pkg,target))continue;
    removePart(pkg,target);removedMedia.push(target);
   }
  }
  pkg.toBytes();return {shapeId:id,partName:part,removedRelationships,removedMedia};
 });
}
