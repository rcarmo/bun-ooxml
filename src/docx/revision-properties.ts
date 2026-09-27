import {attribute,escapeAttribute,type XmlElement} from '../xml/index.ts';
import {appearanceFromNode,type AppearanceKey} from './run-appearance.ts';
import {fontNameFromNode} from './run-font.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
const ORDER=['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath'];
const FLAGS=new Set(['b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','vanish']);
const APPEARANCE:Record<string,AppearanceKey>={color:'color',u:'underline',highlight:'highlight',vertAlign:'verticalAlign'};
const word=(e:XmlElement|undefined,n:string)=>e?.namespaceURI===W&&e.localName===n;
function attributes(e:XmlElement,allowed:string[]):boolean{return Object.keys(e.attributes).every(k=>e.attributeNamespaces[k]===XMLNS||(e.attributeNamespaces[k]===W&&allowed.includes(k.split(':').pop()!)));}
function whitespace(e:XmlElement,xml:string):boolean{if(e.selfClosing)return true;let at=e.openEnd;for(const c of e.children){if(!/^[ \t\r\n]*$/.test(xml.slice(at,c.start)))return false;at=c.end;}return /^[ \t\r\n]*$/.test(xml.slice(at,e.closeStart));}
export function runPropertySetSupported(pr:XmlElement,xml:string,change?:XmlElement):boolean{
 if(!attributes(pr,[])||!whitespace(pr,xml))return false;const seen=new Set<string>();let last=-1;
 for(const node of pr.children){if(node===change){if(node!==pr.children.at(-1))return false;continue;}const name=node.localName,rank=ORDER.indexOf(name);if(node.namespaceURI!==W||rank<last||rank<0||seen.has(name)||node.children.length||!whitespace(node,xml))return false;seen.add(name);last=rank;
  if(FLAGS.has(name)){if(!attributes(node,['val'])||!/^(?:1|0|true|false|on|off)$/i.test(attribute(node,'val',W)??'true'))return false;}
  else if(APPEARANCE[name]){try{appearanceFromNode(APPEARANCE[name]!,node);}catch{return false;}}
  else if(name==='sz'||name==='szCs'){const v=attribute(node,'val',W);if(!attributes(node,['val'])||!v||!/^\d+$/.test(v)||!Number.isSafeInteger(Number(v))||Number(v)<1)return false;}
  else if(name==='rFonts'){try{fontNameFromNode(node);}catch{return false;}}
  else return false;
 }
 for(const pair of [['strike','dstrike'],['caps','smallCaps'],['emboss','imprint'],['emboss','outline'],['emboss','shadow'],['imprint','outline'],['imprint','shadow']])if(pair.every(n=>seen.has(n)))return false;
 return true;
}
/** Bounded direct-text run profile. Both snapshots are checked even on accept. */
export function runPropertyChangeProblem(change:XmlElement,xml:string):string|undefined{
 const pr=change.parent,run=pr?.parent,p=run?.parent;
 if(!word(pr,'rPr')||!word(run,'r')||!word(p,'p')||run!.children[0]!==pr||run!.children.filter(c=>word(c,'rPr')).length!==1)return 'run-property changes require the unique first rPr of a direct paragraph text run';
 for(let n=p;n;n=n.parent)if(n.namespaceURI===W&&(['ins','del','moveFrom','moveTo'].includes(n.localName)||n.localName.endsWith('Change')))return 'nested run-property revisions are unsupported';
 if(!whitespace(run!,xml)||run!.children.some(c=>c!==pr&&(!word(c,'t')||c.children.length||/<!--|<!\[CDATA\[|<\?/.test(xml.slice(c.openEnd,c.closeStart)))))return 'run-property revisions require only plain text runs';
 if(attribute(change,'author',W)===undefined)return 'run-property change requires a qualified author';
 if(!attributes(change,['id','author','date'])||!whitespace(change,xml)||change.children.length!==1||!word(change.children[0],'rPr'))return 'run-property change must contain exactly one undecorated saved rPr snapshot';
 if(!runPropertySetSupported(pr!,xml,change)||!runPropertySetSupported(change.children[0]!,xml))return 'current or saved run properties are ambiguous, decorated, conflicting or unsupported';
 return undefined;
}
/** Keep the saved rPr spelling; transfer bindings from ancestors removed on reject. */
export function restoredRunProperties(change:XmlElement,xml:string):string{
 const saved=change.children[0]!,declarations=new Map<string,string>();for(const e of [change.parent!,change])for(const [k,v]of Object.entries(e.attributes))if(e.attributeNamespaces[k]===XMLNS)declarations.set(k,v);
 for(const k of Object.keys(saved.attributes))declarations.delete(k);
 const raw=xml.slice(saved.start,saved.end),offset=saved.openEnd-saved.start-(saved.selfClosing?2:1),extra=[...declarations].map(([k,v])=>` ${k}="${escapeAttribute(v)}"`).join('');return raw.slice(0,offset)+extra+raw.slice(offset);
}
