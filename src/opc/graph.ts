import {posix} from 'node:path';
import {OpcPackage,relationshipPath,type Relationship} from './package.ts';
import {getContentType,setPartContentType,removePartContentType} from './content-types.ts';
import {parseXml,applyEdits,escapeAttribute,attribute,type XmlElement} from '../xml/index.ts';
import {OoxmlError} from '../errors.ts';
const REL='http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_TYPE='application/vnd.openxmlformats-package.relationships+xml';
function fail(code:string,message:string):never{throw new OoxmlError(code,message);}
function isMeta(name:string):boolean{return name==='[Content_Types].xml'||name.endsWith('.rels');}
function validateOwner(pkg:OpcPackage,owner:string):void{if(owner&&(isMeta(owner)||!pkg.get(owner)))fail('opc-owner-invalid','Relationship owner is not a payload part');}

/** Add an owned part and exact type in one rollback boundary. Opaque bytes are
 * never decoded; unrelated members and default extension mappings remain intact. */
export function addPart(pkg:OpcPackage,name:string,value:Uint8Array|string,contentType:string):void{
 if(isMeta(name))fail('opc-part-reserved','Use relationship/content-type APIs for metadata parts');
 if(pkg.names().some(n=>n.toLowerCase()===name.toLowerCase()))fail('opc-part-collision',`Part already exists: ${name}`);
 pkg.transaction(()=>{pkg.set(name,value);setPartContentType(pkg,name,contentType);pkg.toBytes();});
}

/** No implicit cascade or detach. Callers must remove incoming references first;
 * removing an owner deletes its own .rels but retains every target part. */
export function removePart(pkg:OpcPackage,name:string):void{
 if(isMeta(name))fail('opc-part-reserved','Cannot remove package metadata directly');
 if(!pkg.get(name))fail('opc-part-missing',`Part does not exist: ${name}`);
 if(pkg.mainPart()===name)fail('opc-main-part-required','Cannot remove the main Office part');
 for(const owner of owners(pkg))for(const rel of pkg.relationships(owner))if(!rel.external&&rel.resolved===name)fail('opc-part-referenced',`Part ${name} is referenced by ${owner||'/'}:${rel.id}`);
 pkg.transaction(()=>{
  const rels=relationshipPath(name);
  if(pkg.get(rels)){pkg.delete(rels);removePartContentType(pkg,rels);}
  pkg.delete(name);removePartContentType(pkg,name);pkg.toBytes();
 });
}

/** Internal targets are owner-relative (or package-absolute). External targets
 * are retained as data and never fetched. Reuse an equivalent relationship unless
 * the caller supplies a different explicit ID; generated IDs fill gaps. */
export function addRelationship(pkg:OpcPackage,owner:string,type:string,target:string,options:{id?:string;external?:boolean}={}):Relationship{
 validateOwner(pkg,owner);
 if(!/^[A-Za-z][A-Za-z0-9+.-]*:[^\s\u0000-\u001f]+$/.test(type))fail('opc-relationship-type-invalid','Relationship type must be an absolute URI');
 if(!target||/[\u0000-\u001f]/.test(target))fail('opc-target-invalid','Empty/control-character target');
 const id=options.id;
 if(id!==undefined&&!/^[A-Za-z_][\w.-]*$/.test(id))fail('opc-relationship-id-invalid','Relationship ID must be a nonempty XML name');
 const existing=pkg.relationships(owner);
 const external=options.external===true;
 // Validate candidate using the same package resolver, including fragment and
 // percent-encoding policy. No local URI parser diverges from load validation.
 return pkg.transaction(()=>{
  const path=relationshipPath(owner);
  const beforeRels=pkg.get(path),beforeTypes=pkg.get('[Content_Types].xml')!;
  const candidateId=id??nextRelationshipId(existing);
  if(existing.some(r=>r.id===candidateId))fail('opc-relationship-duplicate',`Relationship ID already exists: ${candidateId}`);
  const xml=pkg.get(path)?pkg.text(path):`<Relationships xmlns="${REL}"/>`;
  const root=parseXml(xml).root;
  const name=qualified(root,'Relationship');
  const child=`<${name} Id="${escapeAttribute(candidateId)}" Type="${escapeAttribute(type)}" Target="${escapeAttribute(target)}"${external?' TargetMode="External"':''}/>`;
  const next=append(xml,root,child);
  pkg.set(path,next);setPartContentType(pkg,path,REL_TYPE);
  const candidate=pkg.relationships(owner).find(r=>r.id===candidateId)!;
  if(!candidate.external&&(!candidate.resolved||!pkg.get(candidate.resolved)||isMeta(candidate.resolved)))fail('opc-relationship-target-missing','Relationship target must be an existing payload part');
  const equivalent=existing.find(r=>r.type===type&&r.external===external&&(external?r.target===target:r.resolved===candidate.resolved&&r.target.split('#')[1]===target.split('#')[1]));
  if(equivalent&&!id){
    // Restore metadata exactly: get-or-add must not dirty an existing package.
    if(beforeRels)pkg.set(path,beforeRels);else pkg.delete(path);
    pkg.set('[Content_Types].xml',beforeTypes);
    return equivalent;
  }
  pkg.toBytes();return {...candidate};
 });
}

/** Remove only when the owner has no officeDocument id/embed/link reference to
 * the relationship. Binary owners refuse because their link grammar is unknown. */
export function removeRelationship(pkg:OpcPackage,owner:string,id:string):void{
 validateOwner(pkg,owner);
 const rel=pkg.relationships(owner).find(r=>r.id===id);if(!rel)fail('opc-relationship-missing',`No relationship ${id}`);
 if(owner){
  const type=getContentType(pkg,owner);
  if(!(type?.endsWith('+xml')||type==='application/xml'||type==='text/xml'))fail('opc-relationship-owner-unsupported','Cannot inspect relationship references in binary owner');
  const xml=parseXml(pkg.text(owner));
  for(const e of xml.elements)for(const local of ['id','embed','link'])if(attribute(e,local,OFFICE)===id)fail('opc-relationship-referenced',`Relationship ${id} is referenced by owner XML`);
 }
 pkg.transaction(()=>{
  const path=relationshipPath(owner),xml=pkg.text(path),document=parseXml(xml);
  const node=document.root.children.find(e=>e.namespaceURI===REL&&e.localName==='Relationship'&&e.attributes.Id===id);
  if(!node)fail('opc-relationship-missing',`No relationship ${id}`);
  pkg.set(path,applyEdits(xml,[{start:node.start,end:node.end,value:''}]));pkg.toBytes();
 });
}

/** Find the first unused numbered name across all members, including orphans. */
export function nextPartName(pkg:OpcPackage,template:string):string{
 if(template.split('%d').length!==2||/%(?!d)/.test(template))fail('opc-part-template-invalid','Part template needs exactly one %d');
 const used=new Set(pkg.names().map(n=>n.toLowerCase()));
 for(let n=1;n<=used.size+1;n++){const name=template.replace('%d',String(n));
  // The content-type helper shares canonical OPC part-name validation without mutation.
  getContentType(pkg,name);
  if(!used.has(name.toLowerCase()))return name;
 }
 fail('opc-part-template-invalid','No name available');
}

/** Iterative DFS follows each internal target once; external targets are data.
 * Unrelated orphan members are preserved but do not appear in this graph view. */
export function walkParts(pkg:OpcPackage,owner=''):string[]{
 validateOwner(pkg,owner);const seen=new Set<string>(owner?[owner]:[]),result:string[]=[];
 const pending=[...pkg.relationships(owner)].reverse();
 while(pending.length){const rel=pending.pop()!;if(rel.external||!rel.resolved||seen.has(rel.resolved))continue;
  seen.add(rel.resolved);result.push(rel.resolved);pending.push(...[...pkg.relationships(rel.resolved)].reverse());
 }return result;
}
function owners(pkg:OpcPackage):string[]{return ['',...pkg.names().filter(n=>!isMeta(n))];}
function nextRelationshipId(rels:Relationship[]):string{const ids=new Set(rels.map(r=>r.id));for(let i=1;i<=ids.size+1;i++){const id=`rId${i}`;if(!ids.has(id))return id;}throw new Error('Unreachable');}
function qualified(root:XmlElement,local:string):string{const colon=root.name.indexOf(':');return colon<0?local:root.name.slice(0,colon+1)+local;}
function append(xml:string,root:XmlElement,child:string):string{
 return root.selfClosing?applyEdits(xml,[{start:root.start,end:root.end,value:xml.slice(root.start,root.openEnd).replace(/\/\s*>$/,()=>`>${child}</${root.name}>`)}]):applyEdits(xml,[{start:root.closeStart,end:root.closeStart,value:child}]);
}
