import {OoxmlError} from '../errors.ts';
import type {OpcPackage} from '../opc/package.ts';
import {getContentType} from '../opc/content-types.ts';
import {parseXml,elements,attribute,type XmlElement} from '../xml/index.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',D='http://schemas.openxmlformats.org/drawingml/2006/diagram',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',DSP='http://schemas.microsoft.com/office/drawing/2008/diagram',DRAW='http://schemas.microsoft.com/office/2007/relationships/diagramDrawing';
export type SmartArtPart={partName:string;contentType:string;byteLength:number};
export type SmartArtRoot=SmartArtPart&{role:'data'|'layout'|'style'|'colors';relationshipId:string;target:string};
export type SmartArtEdge={owner:string;relationshipId:string;type:string;target:string;external:boolean;partName:string|null};
export type SmartArtLimits={dataEditing:boolean;layoutEvaluation:boolean;styleEvaluation:boolean;drawingRegeneration:boolean;reason:string};
export type SmartArtInfo={shapeId:number;name:string;slidePart:string;roots:SmartArtRoot[];parts:SmartArtPart[];edges:SmartArtEdge[];drawingParts:string[];limits:SmartArtLimits};
function fail(message:string):never{throw new OoxmlError('PPTX_SMARTART_UNSUPPORTED',message);}
function one(parent:XmlElement,name:string,ns:string):XmlElement{const rows=parent.children.filter(n=>n.localName===name&&n.namespaceURI===ns);if(rows.length!==1)fail('Expected unique '+name);return rows[0]!;}
const roles=[['data','dm','diagramData','diagramData','dataModel'],['layout','lo','diagramLayout','diagramLayout','layoutDef'],['style','qs','diagramQuickStyle','diagramStyle','styleDef'],['colors','cs','diagramColors','diagramColors','colorsDef']] as const;
function asset(pkg:OpcPackage,partName:string):SmartArtPart{const bytes=pkg.get(partName),contentType=getContentType(pkg,partName);if(!bytes||!contentType)fail('Missing typed SmartArt dependency');return {partName,contentType,byteLength:bytes.length};}
function root(pkg:OpcPackage,part:string,mime:string,name:string,ns=D){if(getContentType(pkg,part)!==mime)fail('Mismatched SmartArt content type');const doc=parseXml(pkg.text(part));if(doc.root.localName!==name||doc.root.namespaceURI!==ns)fail('Mismatched SmartArt XML root');return doc;}
/** Detached bounded package graph only; no layout engine, editing or fetching. */
export function inspectSmartArt(pkg:OpcPackage,part:string):SmartArtInfo[]{
 const doc=parseXml(pkg.text(part));if(doc.root.localName!=='sld'||doc.root.namespaceURI!==P)fail('Expected slide root');const tree=one(one(doc.root,'cSld',P),'spTree',P),ids=new Set<number>();
 for(const n of elements(doc,'cNvPr',P)){const raw=attribute(n,'id'),id=raw&&/^\d+$/.test(raw)?Number(raw):NaN;if(!Number.isInteger(id)||id<1||id>2147483647||ids.has(id))fail('Duplicate or invalid slide identity');ids.add(id);}
 const used=new Set<XmlElement>(),frames:XmlElement[]=[];
 function visit(parent:XmlElement){for(const n of parent.children){if(n.namespaceURI!==P)continue;if(n.localName==='grpSp')visit(n);else if(n.localName==='graphicFrame')frames.push(n);}}visit(tree);
 const result:SmartArtInfo[]=[];
 for(const frame of frames){
  const graphics=frame.children.filter(n=>n.localName==='graphic'&&n.namespaceURI===A);if(!graphics.some(n=>n.children.some(c=>c.namespaceURI===A&&c.localName==='graphicData'&&attribute(c,'uri')===D)))continue;
  const graphic=one(frame,'graphic',A),data=one(graphic,'graphicData',A),leaf=one(data,'relIds',D);if(data.children.length!==1||leaf.children.length||leaf.directText.trim())fail('Ambiguous SmartArt role leaf');used.add(leaf);
  for(const k of Object.keys(leaf.attributes))if(['dm','lo','qs','cs'].includes(k.split(':').at(-1)!)&&leaf.attributeNamespaces[k]!==R)fail('Foreign SmartArt relationship attribute');
  const identity=one(one(frame,'nvGraphicFramePr',P),'cNvPr',P),roots:SmartArtRoot[]=[],relationships=pkg.relationships(part);
  for(const [role,key,type,mime,name]of roles){const id=attribute(leaf,key,R),matches=relationships.filter(r=>r.id===id);if(!id||matches.length!==1)fail('Missing SmartArt role relationship');const r=matches[0]!;if(r.type!==R+'/'+type||r.external||!r.resolved)fail('Wrong/internal SmartArt role relationship');root(pkg,r.resolved,'application/vnd.openxmlformats-officedocument.drawingml.'+mime+'+xml',name);roots.push({...asset(pkg,r.resolved),role,relationshipId:id,target:r.target});}
  const parts=new Map<string,SmartArtPart>(),edges:SmartArtEdge[]=[],drawings=new Set<string>(),pending=roots.map(r=>r.partName).reverse();
  while(pending.length){const owner=pending.pop()!;if(parts.has(owner))continue;if(parts.size>=256)fail('SmartArt dependency part limit');parts.set(owner,asset(pkg,owner));const deps=pkg.relationships(owner);
   if(getContentType(pkg,owner)==='application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml'){
    const dataDoc=parseXml(pkg.text(owner)),metadata=elements(dataDoc,'dataModelExt',DSP);if(metadata.length>1)fail('Ambiguous drawing metadata');for(const m of metadata){
     const id=attribute(m,'relId'),local=deps.filter(r=>r.id===id&&r.type===DRAW&&!r.external),slide=relationships.filter(r=>r.id===id&&r.type===DRAW&&!r.external);
     if(!id||local.length+slide.length!==1)fail('Stale or ambiguous drawing metadata relationship');
     if(slide.length){const r=slide[0]!;if(!r.resolved)fail('Unresolved slide drawing relationship');root(pkg,r.resolved,'application/vnd.ms-office.drawingml.diagramDrawing+xml','drawing',DSP);drawings.add(r.resolved);
      if(!edges.some(e=>e.owner===part&&e.relationshipId===r.id)){if(edges.length>=1024)fail('SmartArt dependency edge limit');edges.push({owner:part,relationshipId:r.id,type:r.type,target:r.target,external:false,partName:r.resolved});}pending.push(r.resolved);
     }
    }
   }
   for(const r of deps){if(edges.length>=1024)fail('SmartArt dependency edge limit');edges.push({owner,relationshipId:r.id,type:r.type,target:r.target,external:r.external,partName:r.resolved??null});if(r.external)continue;if(!r.resolved)fail('Unresolved SmartArt dependency');if(r.type===DRAW){root(pkg,r.resolved,'application/vnd.ms-office.drawingml.diagramDrawing+xml','drawing',DSP);drawings.add(r.resolved);}pending.push(r.resolved);}
  }
  result.push({shapeId:Number(attribute(identity,'id')),name:attribute(identity,'name')??'',slidePart:part,roots,parts:[...parts.values()].sort((a,b)=>a.partName.localeCompare(b.partName)),edges:edges.sort((a,b)=>a.owner.localeCompare(b.owner)||a.relationshipId.localeCompare(b.relationshipId)),drawingParts:[...drawings].sort(),limits:{dataEditing:false,layoutEvaluation:false,styleEvaluation:false,drawingRegeneration:false,reason:'Dependency inspection only; no SmartArt layout/edit engine.'}});
 }
 if(elements(doc,'relIds',D).some(n=>!used.has(n)))fail('Misplaced SmartArt role leaf');return result;
}
