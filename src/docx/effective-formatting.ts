import {OoxmlError} from '../errors.ts';
import {getContentType,type OpcPackage} from '../opc/index.ts';
import {parseXml,attribute,type XmlElement} from '../xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',NS='http://www.w3.org/2000/xmlns/';
const reject=(message:string):never=>{throw new OoxmlError('docx-effective-unsupported',message);};
export type FormattingContribution={source:'implicit'|'document-default'|'paragraph-style'|'direct-run';partName?:string;styleId?:string;operation:'set'|'toggle'|'retain';value:boolean;result:boolean};
export type EffectiveFlag={value:boolean;contributions:FormattingContribution[]};
export type EffectiveRunFormatting={runIndex:number;text:string;paragraphStyleChain:string[];bold:EffectiveFlag;italic:EffectiveFlag};
type Flags={bold?:boolean;italic?:boolean};
function children(node:XmlElement,names:string[]){
 if(node.text.trim()||node.children.some(n=>n.namespaceURI!==W||!names.includes(n.localName)))reject('Unsupported '+node.localName+' content');
 const seen=new Set<string>();for(const child of node.children){if(seen.has(child.localName))reject('Duplicate '+child.localName);seen.add(child.localName);}
}
function attrs(node:XmlElement,names:string[]){for(const key of Object.keys(node.attributes))if(node.attributeNamespaces[key]!==NS&&!(node.attributeNamespaces[key]===W&&names.includes(key.split(':').at(-1)!)))reject('Unsupported attribute on '+node.localName);}
function onOff(value:string|undefined):boolean {if(value===undefined||['1','true','on'].includes(value))return true;if(['0','false','off'].includes(value))return false;return reject('Malformed on/off value');}
function leaf(node:XmlElement){if(node.children.length||node.text.trim())reject('Expected leaf '+node.localName);}
function flags(node:XmlElement|undefined):Flags{
 if(!node)return {};attrs(node,[]);children(node,['b','i']);const out:Flags={};
 for(const child of node.children){attrs(child,['val']);leaf(child);out[child.localName==='b'?'bold':'italic']=onOff(attribute(child,'val',W));}return out;
}
function whitespaceOnlyBetween(xml:string,node:XmlElement){
 let cursor=node.openEnd;for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,child.start)))reject('Mixed or lexical content');cursor=child.end;}
 if(!node.selfClosing&&!/^[ \t\r\n]*$/.test(xml.slice(cursor,node.closeStart)))reject('Mixed or lexical content');
}
function find(node:XmlElement,name:string){return node.children.find(n=>n.namespaceURI===W&&n.localName===name);}
function single(node:XmlElement,name:string):XmlElement|undefined{const nodes=node.children.filter(n=>n.namespaceURI===W&&n.localName===name);if(nodes.length>1)reject('Duplicate '+name);return nodes[0];}
function styleRef(node:XmlElement|undefined):string|undefined{if(!node)return;attrs(node,['val']);leaf(node);const id=attribute(node,'val',W);if(!id||id.trim()!==id)reject('Invalid style reference');return id;}
function paragraphStyle(paragraph:XmlElement):string|undefined{
 const props=single(paragraph,'pPr');if(!props)return;
 if(paragraph.children[0]!==props)reject('Misplaced paragraph properties');attrs(props,[]);children(props,['pStyle']);return styleRef(find(props,'pStyle'));
}
function resolveStyles(pkg:OpcPackage,main:string,selected:string|undefined){
 const relationships=pkg.relationships(main);if(relationships.some(r=>r.type==='http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects'||r.type===R+'/stylesWithEffects'))reject('Styles-with-effects requires independent reconciliation');
 const refs=relationships.filter(r=>r.type===R+'/styles');if(refs.length>1)reject('Ambiguous styles relationship');
 if(!refs.length){if(selected)reject('Missing selected style');return {defaults:{} as Flags,chain:[] as Array<{id:string;flags:Flags}>,partName:undefined};}
 const rel=refs[0]!;if(rel.external||!rel.resolved)reject('Styles must be internal');const partName=rel.resolved!;
 if(getContentType(pkg,partName)!=='application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml')reject('Wrong styles content type');
 const root=parseXml(pkg.text(partName)).root;if(root.namespaceURI!==W||root.localName!=='styles')reject('Invalid styles root');attrs(root,[]);
 if(root.children.some(n=>n.namespaceURI!==W||!['style','docDefaults','latentStyles'].includes(n.localName))||root.text.trim())reject('Unknown styles content');
 const registry=new Map<string,XmlElement>();let defaultId:string|undefined;
 for(const node of root.children.filter(n=>n.localName==='style')){
  attrs(node,['type','styleId','default','customStyle']);const id=attribute(node,'styleId',W),type=attribute(node,'type',W);
  if(!id||id.trim()!==id||registry.has(id)||!['paragraph','character','table','numbering'].includes(type??''))reject('Invalid or duplicate style identity');
  registry.set(id!,node);const defaultValue=attribute(node,'default',W);
  if(defaultValue!==undefined&&onOff(defaultValue)&&type==='paragraph'){if(defaultId!==undefined)reject('Duplicate paragraph default');defaultId=id;}
 }
 const defaults=single(root,'docDefaults');let defaultFlags:Flags={};
 if(defaults){attrs(defaults,[]);children(defaults,['rPrDefault','pPrDefault']);const r=find(defaults,'rPrDefault');if(r){attrs(r,[]);children(r,['rPr']);defaultFlags=flags(find(r,'rPr'));}const p=find(defaults,'pPrDefault');if(p){attrs(p,[]);children(p,['pPr']);const pp=find(p,'pPr');if(pp){attrs(pp,[]);children(pp,[]);}}}
 const chain:Array<{id:string;flags:Flags}>=[],seen=new Set<string>();let id=selected??defaultId;
 while(id!==undefined){
  if(seen.has(id))reject('Cyclic paragraph style ancestry');seen.add(id);const node=registry.get(id);if(!node||attribute(node,'type',W)!=='paragraph')reject('Missing or nonparagraph ancestor');
  children(node!,['name','basedOn','next','link','autoRedefine','hidden','uiPriority','semiHidden','unhideWhenUsed','qFormat','locked','personal','personalCompose','personalReply','rsid','rPr','pPr']);
  if(find(node!,'link'))reject('Linked character style is unsupported');
  const pp=find(node!,'pPr');if(pp){attrs(pp,[]);children(pp,[]);}
  chain.unshift({id,flags:flags(find(node!,'rPr'))});id=styleRef(find(node!,'basedOn'));
 }
 return {defaults:defaultFlags,chain,partName};
}
/** Bold/italic only, for plain direct body runs. No table/numbering/character/CS cascade. */
export function inspectEffectiveFormatting(pkg:OpcPackage,main:string,paragraph:XmlElement):EffectiveRunFormatting[]{
 if(pkg.mainPart()!==main)reject('Document relationship target changed');
 if(paragraph.namespaceURI!==W||paragraph.localName!=='p'||paragraph.parent?.namespaceURI!==W||paragraph.parent.localName!=='body'||paragraph.parent.parent?.namespaceURI!==W||paragraph.parent.parent.localName!=='document')reject('Only direct body paragraphs are supported');
 const xml=pkg.text(main);whitespaceOnlyBetween(xml,paragraph);
 attrs(paragraph,[]);if(paragraph.children.some(n=>n.namespaceURI!==W||!['pPr','r'].includes(n.localName)))reject('Unsupported paragraph content');
 const selected=paragraphStyle(paragraph),styles=resolveStyles(pkg,main,selected);
 return paragraph.children.filter(n=>n.localName==='r').map((run,runIndex)=>{
  whitespaceOnlyBetween(xml,run);attrs(run,[]);if(run.children.some(n=>n.namespaceURI!==W||!['rPr','t'].includes(n.localName)))reject('Only plain direct text runs are supported');
  const props=single(run,'rPr');if(props&&run.children[0]!==props)reject('Misplaced run properties');const direct=flags(props);
  const text=run.children.filter(n=>n.localName==='t').map(n=>{if(n.children.length)reject('Nested text content');for(const key of Object.keys(n.attributes))if(n.attributeNamespaces[key]!==NS&&!(key==='xml:space'&&n.attributeNamespaces[key]==='http://www.w3.org/XML/1998/namespace'))reject('Unsupported text attribute');return n.text;}).join('');
  // Script-dependent bCs/iCs selection is outside this bounded Latin-text slice.
  if(/[^\u0009\u000A\u000D\u0020-\u024F\u0300-\u036F\u2000-\u206F]/u.test(text))reject('Script-dependent formatting is unsupported');
  const effective=(key:keyof Flags):EffectiveFlag=>{
   let value=false;const contributions:FormattingContribution[]=[{source:'implicit',operation:'set',value:false,result:false}];
   const def=styles.defaults[key];if(def!==undefined){value=def;contributions.push({source:'document-default',partName:styles.partName!,operation:'set',value:def,result:value});}
   for(const style of styles.chain){const flag=style.flags[key];if(flag===undefined)continue;if(flag)value=!value;contributions.push({source:'paragraph-style',partName:styles.partName!,styleId:style.id,operation:flag?'toggle':'retain',value:flag,result:value});}
   if(direct[key]!==undefined){value=direct[key]!;contributions.push({source:'direct-run',partName:main,operation:'set',value,result:value});}
   return {value,contributions};
  };
  return {runIndex,text,paragraphStyleChain:styles.chain.map(s=>s.id),bold:effective('bold'),italic:effective('italic')};
 });
}
