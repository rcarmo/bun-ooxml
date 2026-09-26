import {OoxmlError} from '../errors.ts';
import {parseXml,applyEdits,attribute,type XmlElement} from '../xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
const ORDER=['headerReference','footerReference','footnotePr','endnotePr','type','pgSz','pgMar','paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg','textDirection','bidi','rtlGutter','docGrid','printerSettings'];
const margins=['top','right','bottom','left','header','footer','gutter'] as const;
const keys=['width','height','orientation',...margins];
export type PageLayout={width:number;height:number;orientation:'portrait'|'landscape';top:number;right:number;bottom:number;left:number;header:number;footer:number;gutter:number};
const fail=(message:string):never=>{throw new OoxmlError('docx-layout-unsupported',message);};
const word=(n:XmlElement,local:string)=>n.namespaceURI===W&&n.localName===local;
function whitespace(xml:string,node:XmlElement){if(node.selfClosing)return;let at=node.openEnd;for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(at,child.start)))fail('Section lexical barriers are unsupported');at=child.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(at,node.closeStart)))fail('Section lexical barriers are unsupported');}
function geometry(layout:PageLayout){
 for(const key of ['width','height',...margins] as const)if(!Number.isInteger(layout[key])||layout[key]<(key==='width'||key==='height'?1:0)||layout[key]>31680)fail('Geometry must be bounded integer twips');
 if(!['portrait','landscape'].includes(layout.orientation))fail('Invalid orientation');
 if(layout.orientation==='portrait'&&layout.width>layout.height||layout.orientation==='landscape'&&layout.width<layout.height)fail('Size and orientation disagree');
 if(layout.left+layout.right+layout.gutter>=layout.width||layout.top+layout.bottom>=layout.height)fail('Margins leave no content area');
 if(layout.header>layout.height||layout.footer>layout.height)fail('Header/footer distance exceeds page');
}
export function normalizePageLayout(value:PageLayout):PageLayout{
 if(!value||typeof value!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(value)))fail('Layout must be plain data');
 const result:Record<string,unknown>={};for(const key of Reflect.ownKeys(value)){
  if(typeof key!=='string'||!keys.includes(key))fail('Unknown layout property');const d=Object.getOwnPropertyDescriptor(value,key)!;if(!('value'in d))fail('Layout accessors are unsupported');result[key as string]=d.value;
 }
 if(keys.some(key=>!Object.hasOwn(result,key)))fail('All page geometry fields are required');geometry(result as PageLayout);return result as PageLayout;
}
function inspect(xml:string){
 const root=parseXml(xml).root;if(!word(root,'document'))fail('Invalid Word document root');
 const bodies=root.children.filter(n=>word(n,'body'));if(bodies.length!==1)fail('Expected unique body');const body=bodies[0]!,sections=body.children.filter(n=>word(n,'sectPr'));
 if(sections.length!==1||body.children.at(-1)!==sections[0])fail('Expected one final body section');const section=sections[0]!;
 whitespace(xml,section);let last=-1;const seen=new Set<string>();
 for(const node of section.children){const rank=ORDER.indexOf(node.localName);if(node.namespaceURI!==W||rank<0||rank<last||seen.has(node.localName)&&!['headerReference','footerReference'].includes(node.localName))fail('Unknown, duplicate, revised or misordered section metadata');if(['bidi','rtlGutter'].includes(node.localName))fail('Bidi/gutter geometry is unsupported');last=rank;seen.add(node.localName);}
 const size=section.children.find(n=>word(n,'pgSz')),mar=section.children.find(n=>word(n,'pgMar'));if(!size||!mar)fail('Existing page size and margins are required');
 for(const [node,allowed]of [[size!,['w','h','orient','code']],[mar!,margins]] as const){
  if(node.children.length||!node.selfClosing&&xml.slice(node.openEnd,node.closeStart).trim())fail('Page geometry must be empty metadata');
  for(const key of Object.keys(node.attributes))if(node.attributeNamespaces[key]!==XMLNS&&(node.attributeNamespaces[key]!==W||!(allowed as readonly string[]).includes(key.split(':').at(-1)!)))fail('Unsupported or misqualified geometry attribute');
 }
 const number=(node:XmlElement,key:string)=>{const raw=attribute(node,key,W);if(raw===undefined||!/^\d+$/.test(raw))fail('Missing or malformed page geometry');return Number(raw);};
 const layout={width:number(size!,'w'),height:number(size!,'h'),orientation:attribute(size!,'orient',W)??'portrait',...Object.fromEntries(margins.map(key=>[key,number(mar!,key)]))} as PageLayout;
 geometry(layout);return {size:size!,mar:mar!,layout};
}
export function readPageLayout(xml:string):PageLayout{return {...inspect(xml).layout};}
export function replacePageLayout(xml:string,layout:PageLayout):string{
 const {size,mar,layout:old}=inspect(xml);if(keys.every(key=>old[key as keyof PageLayout]===layout[key as keyof PageLayout]))return xml;
 const edits:{start:number;end:number;value:string}[]=[];
 for(const [node,desired]of [[size,{w:String(layout.width),h:String(layout.height),orient:layout.orientation}],[mar,Object.fromEntries(margins.map(k=>[k,String(layout[k])]))]] as const){
  const open=xml.slice(node.start,node.openEnd),missing:Record<string,string>={};
  for(const [key,value]of Object.entries(desired)){
   const current=attribute(node,key,W);if(current===value||key!=='orient'&&current!==undefined&&Number(current)===Number(value)||key==='orient'&&current===undefined&&value==='portrait')continue;
   const qualified=Object.keys(node.attributes).find(k=>node.attributeNamespaces[k]===W&&k.split(':').at(-1)===key);
   if(!qualified){missing[key]=value;continue;}
   const token=[...open.matchAll(/\s+([^\s=/>]+)\s*=\s*("[^"]*"|'[^']*')/g)].find(m=>m[1]===qualified);if(!token)fail('Unresolved geometry attribute offsets');
   const at=node.start+token!.index!;edits.push({start:at,end:at+token![0].length,value:token![0].replace(/(["'])[\s\S]*\1$/,`"${value}"`)});
  }
  if(Object.keys(missing).length){const used=new Set<string>();for(let scope:XmlElement|undefined=node;scope;scope=scope.parent)for(const name of Object.keys(scope.attributes))if(name.startsWith('xmlns:'))used.add(name.slice(6));let prefix='layout';while(used.has(prefix))prefix+='x';const at=node.openEnd-open.match(/\/?\s*>$/)![0].length;edits.push({start:at,end:at,value:` xmlns:${prefix}="${W}"`+Object.entries(missing).map(([k,v])=>` ${prefix}:${k}="${v}"`).join('')});}
 }
 return applyEdits(xml,edits);
}
