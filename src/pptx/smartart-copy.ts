import {posix} from 'node:path';
import {OoxmlError} from '../errors.ts';
import {type OpcPackage,relationshipPath} from '../opc/package.ts';
import {setPartContentType} from '../opc/content-types.ts';
import {addRelationship} from '../opc/graph.ts';
import {inspectSmartArt} from './smartart.ts';
import {shapeAppendSite} from './text-box.ts';
import {parseXml,elements,attribute,applyEdits,escapeAttribute,type XmlElement} from '../xml/index.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',D='http://schemas.openxmlformats.org/drawingml/2006/diagram',DSP='http://schemas.microsoft.com/office/drawing/2008/diagram',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',DRAW='http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';
export type SmartArtCopyReceipt={shapeId:number;partName:string;partMap:Record<string,string>;modelIdMap:Record<string,string>;drawingIdMap:Record<string,number>};
function fail(message:string):never{throw new OoxmlError('PPTX_SMARTART_UNSUPPORTED',message);}
function editsFor(source:string,node:XmlElement,changes:Record<string,string>):{start:number;end:number;value:string}[]{
 const opening=source.slice(node.start,node.openEnd),tokens=/\s+([^\s=/>]+)\s*=\s*(["'])([\s\S]*?)\2/g;
 return Object.entries(changes).map(([key,value])=>{const matches=[...opening.matchAll(tokens)].filter(m=>m[1]===key);if(matches.length!==1)fail('Missing unique lexical copy attribute');const m=matches[0]!,start=node.start+m.index!+m[0].indexOf(m[2]!)+1;return {start,end:start+m[3]!.length,value:escapeAttribute(value)};});
}
const guid=/^\{[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\}$/;
/** Clone only admitted graph types, retarget closed edges and remap instance identities. */
export function copySmartArt(source:OpcPackage,sourcePart:string,id:number,target:OpcPackage,targetPart:string):SmartArtCopyReceipt{
 if(!Number.isInteger(id)||id<1||id>2147483647)fail('Copy requires exact bounded SmartArt ID');
 const info=inspectSmartArt(source,sourcePart).find(r=>r.shapeId===id);if(!info)fail('Missing source SmartArt frame');
 const slideXml=source.text(sourcePart),slideDoc=parseXml(slideXml),frame=elements(slideDoc,'graphicFrame',P).find(n=>n.children.find(c=>c.localName==='nvGraphicFramePr'&&c.namespaceURI===P)?.children.some(c=>c.namespaceURI===P&&c.localName==='cNvPr'&&Number(attribute(c,'id'))===id))!;
 if(frame.parent?.namespaceURI!==P||frame.parent.localName!=='spTree')fail('Grouped SmartArt copy is unsupported');
 let site:ReturnType<typeof shapeAppendSite>;try{shapeAppendSite(slideXml);site=shapeAppendSite(target.text(targetPart));}catch(error){if(error instanceof OoxmlError&&error.code==='PPTX_TEXT_BOX_UNSUPPORTED')fail(error.message);throw error;}
 const roles=['diagramData','diagramLayout','diagramQuickStyle','diagramColors','image','hyperlink'];
 for(const e of info.edges){
  if(e.target.includes('#')||!(e.type===DRAW||roles.some(t=>e.type===R+'/'+t))||e.external&&![R+'/image',R+'/hyperlink'].includes(e.type))fail('Unsupported SmartArt dependency edge');
  if(!e.external){const dependency=info.parts.find(p=>p.partName===e.partName)!;const expected:Record<string,string>={[R+'/diagramData']:'diagramData',[R+'/diagramLayout']:'diagramLayout',[R+'/diagramQuickStyle']:'diagramStyle',[R+'/diagramColors']:'diagramColors'};
   if(e.type===R+'/image'&&!dependency.contentType.startsWith('image/')||e.type===R+'/hyperlink'||expected[e.type]&&dependency.contentType!=='application/vnd.openxmlformats-officedocument.drawingml.'+expected[e.type]+'+xml')fail('Dependency role/type mismatch');
  }
 }
 const partMap:Record<string,string>={},used=new Set(target.names().map(n=>n.toLowerCase())),modelIdMap:Record<string,string>={},drawingIdMap:Record<string,number>={},values=new Map<string,Uint8Array|string>();
 function name(template:string){for(let i=1;i<=used.size+1;i++){const n=template.replace('%d',String(i));if(!used.has(n.toLowerCase())&&!used.has(relationshipPath(n).toLowerCase())){used.add(n.toLowerCase());used.add(relationshipPath(n).toLowerCase());return n;}}return fail('No copy part name available');}
 const diagramMime=/^application\/vnd\.openxmlformats-officedocument\.drawingml\.diagram(?:Data|Layout|Style|Colors)\+xml$/;
 for(const p of info.parts){if(p.contentType.startsWith('image/')){if(source.relationships(p.partName).length)fail('Image dependency owns unsupported outgoing edges');const ext=posix.extname(p.partName);if(!/^\.[a-zA-Z0-9]+$/.test(ext))fail('Unsupported image extension');partMap[p.partName]=name('ppt/media/smartArt%d'+ext);}else if(diagramMime.test(p.contentType)||p.contentType==='application/vnd.ms-office.drawingml.diagramDrawing+xml')partMap[p.partName]=name('ppt/diagrams/smartArt%d.xml');else fail('Unsupported SmartArt dependency content type');}
 const xmlParts=new Map<string,ReturnType<typeof parseXml>>(),xmlText=new Map<string,string>();
 for(const p of info.parts){if(p.contentType.startsWith('image/')){values.set(p.partName,source.get(p.partName)!);continue;}const text=source.text(p.partName),doc=parseXml(text);xmlParts.set(p.partName,doc);xmlText.set(p.partName,text);for(const n of [...elements(doc,'pt',D),...elements(doc,'cxn',D)]){const old=attribute(n,'modelId');if(!old||!guid.test(old)||Object.hasOwn(modelIdMap,old))fail('Duplicate or invalid model identity');modelIdMap[old]='';}}
 const reserved=new Set(Object.keys(modelIdMap).map(v=>v.toUpperCase()));for(const old of Object.keys(modelIdMap)){let v:string;do{v='{'+crypto.randomUUID().toUpperCase()+'}';}while(reserved.has(v));reserved.add(v);modelIdMap[old]=v;}
 const knownReferences=['modelId','srcId','destId','parTransId','sibTransId','presId','presAssocID'];
 for(const [part,doc]of xmlParts){const text=xmlText.get(part)!,edits:{start:number;end:number;value:string}[]=[];const localIDs=new Map<string,number>(),drawNodes=elements(doc,'cNvPr',DSP);
  if(drawNodes.length){let maximum=0;for(const n of drawNodes){const raw=attribute(n,'id');if(!raw||!/^\d+$/.test(raw)||Number(raw)<1||Number(raw)>2147483647||localIDs.has(raw))fail('Invalid drawing shape identity');localIDs.set(raw,0);maximum=Math.max(maximum,Number(raw));}if(maximum+drawNodes.length>2147483647)fail('Drawing identity exhaustion');for(const n of drawNodes){const old=attribute(n,'id')!,next=++maximum;localIDs.set(old,next);drawingIdMap[part+'#'+old]=next;}}
  for(const n of doc.elements){const changes:Record<string,string>={};
   for(const [key,v]of Object.entries(n.attributes)){const local=key.split(':').at(-1)!;if(knownReferences.includes(local)&&n.attributeNamespaces[key]===''){if(v==='')continue;if(!Object.hasOwn(modelIdMap,v))fail('Dangling model reference');changes[key]=modelIdMap[v]!;}else if(guid.test(v)&&local!=='uniqueId'&&local!=='uri')fail('Unknown GUID reference grammar');}
   if(n.namespaceURI===DSP&&n.localName==='cNvPr')changes.id=String(localIDs.get(attribute(n,'id')!)!);
   if(n.namespaceURI===A&&['stCxn','endCxn'].includes(n.localName)){const old=attribute(n,'id')!;if(!localIDs.has(old))fail('Dangling drawing attachment');changes.id=String(localIDs.get(old)!);}
   if(Object.keys(changes).length)edits.push(...editsFor(text,n,changes));
  }
  values.set(part,applyEdits(text,edits));
 }
 const targetXml=target.text(targetPart),frameText=slideXml.slice(frame.start,frame.end);
 return target.transaction(()=>{
  // Install every payload before creating cycle-bearing relationship members.
  for(const p of info.parts){const mapped=partMap[p.partName]!;target.set(mapped,values.get(p.partName)!);setPartContentType(target,mapped,p.contentType);}
  for(const owner of info.parts){const path=relationshipPath(owner.partName);if(!source.get(path))continue;const text=source.text(path),doc=parseXml(text),edits=[];for(const n of doc.root.children){const edge=info.edges.find(e=>e.owner===owner.partName&&e.relationshipId===attribute(n,'Id'))!;if(!edge)fail('Uninspected dependency relationship');if(!edge.external)edits.push(...editsFor(text,n,{Target:posix.relative(posix.dirname(partMap[owner.partName]!),partMap[edge.partName!]!)}));}const mappedPath=relationshipPath(partMap[owner.partName]!);target.set(mappedPath,applyEdits(text,edits));setPartContentType(target,mappedPath,'application/vnd.openxmlformats-package.relationships+xml');}
  const roleMap=new Map<string,string>();for(const r of info.roots){const old=source.relationships(sourcePart).find(e=>e.id===r.relationshipId)!;const rel=addRelationship(target,targetPart,old.type,posix.relative(posix.dirname(targetPart),partMap[r.partName]!));roleMap.set(r.relationshipId,rel.id);}
  const bound=frameTextWithBindings(frameText,frame),fd=parseXml(bound),identity=elements(fd,'cNvPr',P)[0]!,leaf=elements(fd,'relIds',D)[0]!,changes:Record<string,string>={};for(const[k,v]of Object.entries(leaf.attributes))if(leaf.attributeNamespaces[k]===R){const mapped=roleMap.get(v);if(!mapped)fail('Unsupported frame relationship');changes[k]=mapped;}
  const copied=applyEdits(bound,[...editsFor(bound,identity,{id:String(site.shapeId)}),...editsFor(bound,leaf,changes)]);
  target.set(targetPart,applyEdits(targetXml,[{start:site.at,end:site.at,value:copied}]));target.toBytes();return {shapeId:site.shapeId,partName:targetPart,partMap:{...partMap},modelIdMap:{...modelIdMap},drawingIdMap:{...drawingIdMap}};
 });
}
function frameTextWithBindings(text:string,frame:XmlElement):string{
 const ancestors:XmlElement[]=[];for(let n:XmlElement|undefined=frame;n;n=n.parent)ancestors.unshift(n);
 const bindings:Record<string,string>={};for(const n of ancestors)for(const[k,v]of Object.entries(n.attributes))if(k==='xmlns'||k.startsWith('xmlns:')||k.startsWith('xml:'))bindings[k]=v;
 const end=frame.openEnd-frame.start-1;let additions='';for(const[k,v]of Object.entries(bindings))if(!Object.hasOwn(frame.attributes,k))additions+=' '+k+'="'+escapeAttribute(v)+'"';return text.slice(0,end)+additions+text.slice(end);
}
