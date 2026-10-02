import {OoxmlError} from '../errors.ts';
import type {OpcPackage} from '../opc/index.ts';
import {attribute,parseXml,type XmlElement} from '../xml/index.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export type SlideVisibilityObservation=Readonly<{ordinal:number;slideId:string;relationshipId:string;part:string;show:string|null;namespacedShow:string|null;declaredHidden:boolean;applicationConfirmedHidden:false}>;
const refuse=(why:string):never=>{throw new OoxmlError('PPTX_VISIBILITY_INSPECTION_UNSUPPORTED',why);};
const direct=(e:XmlElement,name:string)=>e.children.filter(c=>c.namespaceURI===P&&c.localName===name);
/** Read unqualified schema visibility and the retained namespaced marker separately.
 * No editing, external fetching or application/rendering confirmation is performed.
 */
export function inspectSlideVisibility(pkg:OpcPackage):readonly SlideVisibilityObservation[] {
 const main=pkg.mainPart(),root=parseXml(pkg.text(main)).root;
 if(root.namespaceURI!==P||root.localName!=='presentation')refuse('Expected a presentation root');
 const lists=direct(root,'sldIdLst');if(lists.length!==1)refuse('Expected one direct slide list');
 const entries=direct(lists[0]!,'sldId');if(entries.length!==lists[0]!.children.length)refuse('Unsupported slide list ownership');
 const rels=pkg.relationships(main),byId=new Map(rels.map(r=>[r.id,r]));if(byId.size!==rels.length)refuse('Duplicate relationship identity');
 const ids=new Set<string>(),rids=new Set<string>(),parts=new Set<string>(),rows:SlideVisibilityObservation[]=[];
 for(const [index,e]of entries.entries()){
  const id=attribute(e,'id'),rid=attribute(e,'id',R);
  if(!id||!/^\d+$/.test(id)||!Number.isSafeInteger(Number(id))||Number(id)<1||!rid||ids.has(id)||rids.has(rid))refuse('Missing or duplicate slide identity');
  const rel=byId.get(rid!);if(!rel||rel.external||rel.type!==R+'/slide'||!rel.resolved||parts.has(rel.resolved))refuse('Slide relationship ownership');
  const part=rel!.resolved!,slide=parseXml(pkg.text(part)).root;
  if(slide.namespaceURI!==P||slide.localName!=='sld')refuse('Expected a slide root');
  for(const key of Object.keys(slide.attributes))if(key.split(':').at(-1)==='show'&&!['',P].includes(slide.attributeNamespaces[key]!))refuse('Foreign show attribute');
  const show=attribute(slide,'show')??null,namespacedShow=attribute(slide,'show',P)??null;
  for(const value of [show,namespacedShow])if(value!==null&&!['0','1','true','false'].includes(value))refuse('Invalid visibility value');
  rows.push(Object.freeze({ordinal:index+1,slideId:id!,relationshipId:rid!,part,show,namespacedShow,declaredHidden:show==='0'||show==='false',applicationConfirmedHidden:false}));
  ids.add(id!);rids.add(rid!);parts.add(part);
 }
 return Object.freeze(rows);
}
