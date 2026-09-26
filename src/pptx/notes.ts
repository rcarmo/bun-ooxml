import {OoxmlError} from '../errors.ts';
import {OpcPackage,relationshipPath,sameBytes} from '../opc/package.ts';
import {getContentType} from '../opc/content-types.ts';
import {applyEdits,attribute,elements,escapeText,parseXml,type XmlElement} from '../xml/index.ts';
const P='http://schemas.openxmlformats.org/presentationml/2006/main',A='http://schemas.openxmlformats.org/drawingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide';
const NOTES_TYPE='application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml';
export interface NotesAnchor {readonly kind:'pptx-notes';readonly part:string;readonly text:string;}
type Snapshot={pkg:OpcPackage;slide:string;part:string;epoch:number;bytes:Map<string,Uint8Array|undefined>};
const anchors=new WeakMap<object,Snapshot>(),epochs=new WeakMap<OpcPackage,Map<string,number>>();
const fail=(message:string):never=>{throw new OoxmlError('PPTX_NOTES_UNSUPPORTED',message);};
const is=(n:XmlElement,uri:string,name:string)=>n.namespaceURI===uri&&n.localName===name;
function only(parent:XmlElement,uri:string,name:string):XmlElement{const rows=parent.children.filter(n=>is(n,uri,name));if(rows.length!==1)fail(`Expected exactly one ${name}`);return rows[0]!;}
function gaps(node:XmlElement,xml:string){if(node.selfClosing)return;let at=node.openEnd;for(const n of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(at,n.start)))fail('Unsupported mixed text or markup');at=n.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(at,node.closeStart)))fail('Unsupported mixed text or markup');}
function attrs(node:XmlElement,allowed:string[]){for(const key of Object.keys(node.attributes)){const uri=node.attributeNamespaces[key];if(uri==='http://www.w3.org/2000/xmlns/')continue;if(key==='xml:space'&&uri==='http://www.w3.org/XML/1998/namespace'&&allowed.includes(key))continue;if(uri!==''||!allowed.includes(key))fail('Unsupported notes attribute');}}
// Deliberately bounded template-property subset, not a DrawingML schema validator.
const properties:Record<string,{attrs:string[];children?:string[];required?:string}>={
 pPr:{attrs:['algn','marL','marR','lvl','indent','rtl','fontAlgn','defTabSz','eaLnBrk','latinLnBrk','hangingPunct'],children:['buChar','defRPr']},
 rPr:{attrs:['b','i','lang','altLang','sz','u','strike','kern','cap','spc','baseline','dirty'],children:['solidFill','latin']},
 defRPr:{attrs:['b','i','lang','sz'],children:['solidFill','latin']},
 endParaRPr:{attrs:['b','i','lang','altLang','sz'],children:['solidFill','latin']},
 solidFill:{attrs:[],children:['srgbClr']},srgbClr:{attrs:['val'],required:'val'},latin:{attrs:['typeface','panose','pitchFamily','charset'],required:'typeface'},buChar:{attrs:['char'],required:'char'},
};
function property(node:XmlElement,xml:string){if(node.namespaceURI!==A)fail('Unsupported notes template property');const rule=properties[node.localName]??fail('Unsupported notes template property');attrs(node,rule.attrs);gaps(node,xml);if(rule.required&&attribute(node,rule.required)===undefined)fail('Missing template value');let last=-1;const seen=new Set<string>();for(const child of node.children){const rank=rule.children?.indexOf(child.localName)??-1;if(rank<0||rank<last||seen.has(child.localName))fail('Unsupported template property order or choice');last=rank;seen.add(child.localName);property(child,xml);}if(node.localName==='solidFill'&&node.children.length!==1)fail('Ambiguous solid fill');}
function all(node:XmlElement):XmlElement[]{return [node,...node.children.flatMap(all)];}
function epoch(pkg:OpcPackage,part:string){return epochs.get(pkg)?.get(part)??0;}
function model(pkg:OpcPackage,slide:string){
 const main=pkg.mainPart();if(elements(parseXml(pkg.text(main)),'modifyVerifier',P).length)throw new OoxmlError('PPTX_PROTECTED','Presentation modification protection refuses notes edits');
 const relationships=pkg.relationships(slide).filter(r=>r.type===R);
 if(!relationships.length)throw new OoxmlError('PPTX_NOTES_MISSING','Slide has no existing notes part');
 if(relationships.length!==1||relationships[0]!.external||!relationships[0]!.resolved||relationships[0]!.target.includes('#'))fail('Ambiguous or external notes relationship');
 const part=relationships[0]!.resolved!;if(getContentType(pkg,part)!==NOTES_TYPE)fail('Unexpected notes content type');
 let inbound=0;for(const path of pkg.names().filter(n=>n.endsWith('.rels'))){const match=/^(.*\/)?_rels\/([^/]+)\.rels$/.exec(path);const owner=path==='_rels/.rels'?'':match?(match[1]??'')+match[2]:fail('Unknown relationship owner');for(const rel of pkg.relationships(owner))if(!rel.external&&rel.resolved===part)inbound++;}if(inbound!==1)fail('Notes part has shared or ambiguous ownership');
 const xml=pkg.text(part),doc=parseXml(xml);if(!is(doc.root,P,'notes'))fail('Unexpected notes root');
 if(doc.elements.some(n=>n.localName==='extLst'||![P,A].includes(n.namespaceURI)))fail('Unknown notes namespace or extension');
 if(doc.root.children.some(n=>n.namespaceURI!==P||!['cSld','clrMapOvr'].includes(n.localName)))fail('Unsupported notes root structure');
 const common=only(doc.root,P,'cSld');if(common.children.some(n=>!is(n,P,'spTree')))fail('Unsupported common-slide structure');
 const tree=only(common,P,'spTree');
 const bodies=tree.children.filter(n=>is(n,P,'sp')&&n.children.filter(c=>is(c,P,'nvSpPr')).flatMap(c=>c.children.filter(v=>is(v,P,'nvPr'))).flatMap(v=>v.children.filter(ph=>is(ph,P,'ph'))).some(ph=>attribute(ph,'type')==='body'));
 if(bodies.length!==1)fail('Expected one direct body placeholder');
 const shape=bodies[0]!,nv=only(only(shape,P,'nvSpPr'),P,'nvPr');if(nv.children.filter(c=>is(c,P,'ph')).length!==1)fail('Ambiguous placeholder metadata');
 for(const lock of elements(shape,'spLocks',A))for(const name of Object.keys(lock.attributes))if(lock.attributeNamespaces[name]!=='http://www.w3.org/2000/xmlns/'&&name!=='noGrp')fail('Notes text locks refuse editing');
 if(shape.children.some(n=>n.namespaceURI!==P||!['nvSpPr','spPr','txBody','style'].includes(n.localName)))fail('Unsupported body shape child');
 const body=only(shape,P,'txBody');gaps(body,xml);attrs(body,[]);
 const bodyPr=only(body,A,'bodyPr');attrs(bodyPr,[]);gaps(bodyPr,xml);if(bodyPr.children.length)fail('Unsupported text body properties');
 for(const list of body.children.filter(n=>is(n,A,'lstStyle'))){attrs(list,[]);gaps(list,xml);if(list.children.length)fail('Unsupported list-style inheritance');}
 if(all(body).some(n=>n.namespaceURI!==A&&n!==body||n.localName==='extLst'))fail('Extended notes structure is unsupported');
 if(body.children.some(n=>n.namespaceURI!==A||!['bodyPr','lstStyle','p'].includes(n.localName)))fail('Unsupported notes text body');
 if(body.children.filter(n=>is(n,A,'bodyPr')).length!==1||body.children.filter(n=>is(n,A,'lstStyle')).length>1)fail('Ambiguous text body properties');
 const paragraphs=body.children.filter(n=>is(n,A,'p'));if(!paragraphs.length)fail('Missing notes paragraph');
 let last=-1;for(const node of body.children){const rank=node.localName==='bodyPr'?0:node.localName==='lstStyle'?1:2;if(rank<last)fail('Out-of-order text body');last=rank;}
 const texts:string[]=[];
 for(const p of paragraphs){gaps(p,xml);attrs(p,[]);let rank=-1;const seen=new Set<string>();let text='';
  for(const node of p.children){const r=is(node,A,'pPr')?0:is(node,A,'r')?1:is(node,A,'endParaRPr')?2:-1;if(r<0||r<rank||r!==1&&seen.has(node.localName))fail('Unsupported notes paragraph structure');rank=r;seen.add(node.localName);
   if(r!==1){property(node,xml);continue;}
   gaps(node,xml);attrs(node,[]);if(node.children.some(c=>!is(c,A,'rPr')&&!is(c,A,'t')))fail('Unsupported run topology');
   const properties=node.children.filter(c=>is(c,A,'rPr'));if(properties.length>1||properties.length&&node.children[0]!==properties[0])fail('Ambiguous run properties');
   if(properties[0])property(properties[0],xml);
   const leaf=only(node,A,'t');if(leaf.children.length)fail('Nested notes text');attrs(leaf,['xml:space']);const space=attribute(leaf,'space','http://www.w3.org/XML/1998/namespace');if(space!==undefined&&space!=='preserve'&&space!=='default')fail('Invalid whitespace policy');if(!leaf.selfClosing&&xml.slice(leaf.openEnd,leaf.closeStart).includes('<'))fail('Text leaf has markup barriers');text+=leaf.text;
  }
  texts.push(text);
 }
 return {part,xml,paragraphs,text:texts.join('\n'),main};
}
export function inspectNotes(pkg:OpcPackage,slide:string):NotesAnchor {
 const m=model(pkg,slide),anchor=Object.freeze({kind:'pptx-notes' as const,part:m.part,text:m.text});
 const paths=[...new Set(['[Content_Types].xml',m.main,slide,m.part,...pkg.names().filter(n=>n.endsWith('.rels'))])];
 anchors.set(anchor,{pkg,slide,part:m.part,epoch:epoch(pkg,m.part),bytes:new Map(paths.map(p=>[p,pkg.get(p)]))});return anchor;
}
export function replaceNotes(pkg:OpcPackage,slide:string,anchor:NotesAnchor,text:string):{changedParts:string[]}{
 if(typeof text!=='string'||/[\r\t]/.test(text))throw new OoxmlError('PPTX_NOTES_TEXT_INVALID','Notes require text with LF-only line separators and no tabs');
 escapeText(text); // Includes surrogate/control-character checks before any mutation.
 const state=anchors.get(anchor);if(!state||state.pkg!==pkg||state.slide!==slide||state.epoch!==epoch(pkg,state.part))throw new OoxmlError('PPTX_NOTES_STALE','Notes target is foreign or stale');
 if(pkg.names().filter(n=>n.endsWith('.rels')).some(n=>!state.bytes.has(n)))throw new OoxmlError('PPTX_NOTES_STALE','Notes relationship graph changed');
 for(const [path,bytes]of state.bytes){const now=pkg.get(path);if(bytes===undefined?now!==undefined:now===undefined||!sameBytes(bytes,now))throw new OoxmlError('PPTX_NOTES_STALE','Notes target has stale package bytes');}
 const m=model(pkg,slide);if(m.text===text)return {changedParts:[]};
 const first=m.paragraphs[0]!,runs=first.children.filter(n=>is(n,A,'r'));
 let next:string;
 if(m.paragraphs.length===1&&!text.includes('\n')&&runs.length===1){
  const leaf=only(runs[0]!,A,'t'),space=attribute(leaf,'space','http://www.w3.org/XML/1998/namespace');
  if(!leaf.selfClosing&&(!/^ | $/.test(text)||space==='preserve'))next=applyEdits(m.xml,[{start:leaf.openEnd,end:leaf.closeStart,value:escapeText(text)}]);
  else next=applyEdits(m.xml,[{start:leaf.start,end:leaf.end,value:`<a:t xmlns:a="${A}" xml:space="preserve">${escapeText(text)}</a:t>`}]);
 }else{
  const open=(node:XmlElement)=>m.xml.slice(node.start,node.openEnd).replace(/\/[ \t\r\n]*>$/, '>');
  const pPr=first.children.find(n=>is(n,A,'pPr')),end=first.children.find(n=>is(n,A,'endParaRPr')),run=runs[0],rPr=run?.children.find(n=>is(n,A,'rPr'));
  const fragment=(n:XmlElement|undefined)=>n?m.xml.slice(n.start,n.end):'';
  const paragraphs=text.split('\n').map(line=>`${open(first)}${fragment(pPr)}${run?open(run):`<a:r xmlns:a="${A}">`}${fragment(rPr)}<a:t xmlns:a="${A}" xml:space="preserve">${escapeText(line)}</a:t></${run?.name??'a:r'}>${fragment(end)}</${first.name}>`).join('');
  next=applyEdits(m.xml,[{start:first.start,end:m.paragraphs.at(-1)!.end,value:paragraphs}]);
 }
 pkg.transaction(()=>{pkg.set(m.part,next);if(model(pkg,slide).text!==text)fail('Notes readback differs after editing');pkg.toBytes();});
 const map=epochs.get(pkg)??new Map<string,number>();map.set(m.part,state.epoch+1);epochs.set(pkg,map);return {changedParts:[m.part]};
}
