import {OoxmlError} from '../errors.ts';
import {applyEdits,attribute,type XmlElement} from '../xml/index.ts';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS='http://www.w3.org/2000/xmlns/';
const ORDER=['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath'];
type Edit={start:number;end:number;value:string};
const FLAG_NAMES={bold:'b',italic:'i',caps:'caps',smallCaps:'smallCaps',strike:'strike',doubleStrike:'dstrike',outline:'outline',shadow:'shadow',emboss:'emboss',imprint:'imprint',vanish:'vanish'} as const;
type FlagKey=keyof typeof FLAG_NAMES;
export type DirectRunFlags={ [K in FlagKey]: boolean|null };
export type DirectRunPatch=Partial<DirectRunFlags>&{fontSizePt?:number|null};
const FLAG_KEYS=Object.keys(FLAG_NAMES) as FlagKey[];
const PATCH_KEYS=[...FLAG_KEYS,'fontSizePt'] as const;
const CONFLICTS=[['strike','dstrike'],['caps','smallCaps'],['emboss','imprint'],['emboss','outline'],['imprint','outline'],['emboss','shadow'],['imprint','shadow']] as const;
function fail(message:string):never{throw new OoxmlError('docx-format-unsupported',message);}
function normalizePatch(patch:DirectRunPatch):DirectRunPatch {
 if(!patch||typeof patch!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(patch)))throw new OoxmlError('docx-format-argument','Expected a plain run-formatting patch');
 const result:DirectRunPatch={};
 for(const key of Reflect.ownKeys(patch)){
  if(typeof key!=='string'||!PATCH_KEYS.includes(key as typeof PATCH_KEYS[number]))throw new OoxmlError('docx-format-argument','Unknown run-formatting property');
  const descriptor=Object.getOwnPropertyDescriptor(patch,key)!;
  if(!('value' in descriptor))throw new OoxmlError('docx-format-argument','Run-formatting properties must be plain values');
  const value=descriptor.value;
  if(value===undefined)continue;
  if(value!==null){if(key==='fontSizePt')validateFontSize(value);else if(typeof value!=='boolean')throw new OoxmlError('docx-format-argument','Run effects require boolean or null');}
  Object.assign(result,{[key]:value});
 }
 return result;
}
function assertCompatible(names:Set<string>):void {
 for(const [a,b] of CONFLICTS)if(names.has(a)&&names.has(b))fail(`Incompatible direct run properties: ${a} and ${b}; remove one with null`);
}
const word=(node:XmlElement,name:string)=>node.namespaceURI===W&&node.localName===name;

/** Preserve source outside selected Boolean/size elements or a missing rPr. */
export function formatRunProperties(xml:string,paragraph:XmlElement,patch:DirectRunPatch):{xml:string;changedRuns:number}{
  patch=normalizePatch(patch);
  const requested=PATCH_KEYS.filter(k=>patch[k]!==undefined).sort((a,b)=>ORDER.indexOf(a==='fontSizePt'?'sz':FLAG_NAMES[a])-ORDER.indexOf(b==='fontSizePt'?'sz':FLAG_NAMES[b]));
  const edits:Edit[]=[];let changedRuns=0;
  assertWhitespaceGaps(paragraph,xml);
  const properties=paragraph.children.filter(c=>word(c,'pPr'));
  if(properties.length>1||properties.length&&paragraph.children[0]!==properties[0])fail('Paragraph properties must be unique and first');
  for(const p of properties)if(hasRevision(p))fail('Paragraph property revisions require a review API');
  for(const run of paragraph.children){
    if(word(run,'pPr'))continue;
    if(!word(run,'r'))fail('Formatting requires direct plain text runs');
    assertWhitespaceGaps(run,xml);
    const rprs=run.children.filter(c=>word(c,'rPr'));
    if(rprs.length>1||rprs.length&&run.children[0]!==rprs[0])fail('Run properties must be unique and first');
    if(run.children.some(c=>!word(c,'rPr')&&!word(c,'t')))fail('Fields, drawings and other run content refuse');
    for(const t of run.children.filter(c=>word(c,'t')))if(t.children.length)fail('Text leaves cannot contain child elements');
    const pr=rprs[0];
    if(pr){
      assertWhitespaceGaps(pr,xml);const seen=new Set<string>();let last=-1;
      for(const node of pr.children){
        const rank=ORDER.indexOf(node.localName);
        if(node.namespaceURI!==W||rank<0||seen.has(node.localName)||rank<last||node.children.length)fail('Ambiguous, revised or unsupported run properties');
        seen.add(node.localName);last=rank;
        if(Object.values(FLAG_NAMES).includes(node.localName as typeof FLAG_NAMES[FlagKey]))flagValue(node);
        if(requested.includes('fontSizePt')&&node.localName==='sz')sizeValue(node);
        if(!node.selfClosing&&xml.slice(node.openEnd,node.closeStart).trim())fail('Run properties must not contain text or markup');
      }
    }
    const present=new Set(pr?.children.map(n=>n.localName)??[]);
    assertCompatible(present);
    for(const key of requested){const name=key==='fontSizePt'?'sz':FLAG_NAMES[key];if(patch[key]===null)present.delete(name);else present.add(name);}
    assertCompatible(present);
    let changed=false;const insertions=new Map<number,string[]>();
    for(const key of requested){
      const name=key==='fontSizePt'?'sz':FLAG_NAMES[key],value=patch[key]!,node=pr?.children.find(c=>word(c,name));
      const current=node?(key==='fontSizePt'?sizeValue(node):flagValue(node)):undefined;
      if(value===null){if(node){edits.push({start:node.start,end:node.end,value:''});changed=true;}continue;}
      if(node&&current===value)continue;
      const encoded=key==='fontSizePt'?String((value as number)*2):value?'1':'0';
      const property=`<w:${name} xmlns:w="${W}" w:val="${encoded}"/>`;
      if(node)edits.push({start:node.start,end:node.end,value:property});
      else{
        const following=pr?.children.find(c=>ORDER.indexOf(c.localName)>ORDER.indexOf(name));
        const position=following?.start??pr?.closeStart??run.openEnd;
        const group=insertions.get(position)??[];group.push(property);insertions.set(position,group);
      }
      changed=true;
    }
    if(!changed)continue;
    if(!pr){
      // A self-closing empty run has no text or visible formatting target.
      if(run.selfClosing)fail('Formatting an empty self-closing run is unsupported');
      const value=[...insertions.values()].flat().join('');
      edits.push({start:run.openEnd,end:run.openEnd,value:`<w:rPr xmlns:w="${W}">${value}</w:rPr>`});
    }else if(pr.selfClosing){
      const tail=xml.slice(pr.start,pr.end).match(/\/[ \t\r\n]*>$/);
      if(!tail)fail('Invalid property boundary');
      const open=xml.slice(pr.start,pr.end).replace(/\/[ \t\r\n]*>$/,'>');
      edits.push({start:pr.start,end:pr.end,value:open+[...insertions.values()].flat().join('')+`</${pr.name}>`});
    }else for(const [start,values]of insertions)edits.push({start,end:start,value:values.join('')});
    changedRuns++;
  }
  return {xml:edits.length?applyEdits(xml,edits):xml,changedRuns};
}
/** Direct Boolean overrides in run order; no style inheritance or defaults. */
export function directRunFlags(xml:string,paragraph:XmlElement):DirectRunFlags[]{
 formatRunProperties(xml,paragraph,{});
 return paragraph.children.filter(r=>word(r,'r')).map(run=>{
  const pr=run.children.find(c=>word(c,'rPr'));
  return Object.fromEntries(FLAG_KEYS.map(key=>{const node=pr?.children.find(c=>word(c,FLAG_NAMES[key]));return [key,node?flagValue(node):null];})) as DirectRunFlags;
 });
}
/** Direct non-complex-script sizes in paragraph run order; null means absent. */
export function directFontSizes(xml:string,paragraph:XmlElement):Array<number|null>{
  // Reuse the editing preflight without requesting changes.
  formatRunProperties(xml,paragraph,{});
  return paragraph.children.filter(r=>word(r,'r')).map(run=>{
    const node=run.children.find(c=>word(c,'rPr'))?.children.find(c=>word(c,'sz'));
    return node?sizeValue(node):null;
  });
}
function validateFontSize(value:number):void {
  // Bounded editor policy: positive integer half-points representable without rounding.
  if(typeof value!=='number'||!Number.isFinite(value)||!Number.isSafeInteger(value*2)||value<=0)throw new OoxmlError('docx-format-argument','fontSizePt must be positive and exactly representable in half-points');
}
function sizeValue(node:XmlElement):number{
  for(const key of Object.keys(node.attributes))if(node.attributeNamespaces[key]!==XMLNS&&!(key.split(':').at(-1)==='val'&&node.attributeNamespaces[key]===W))fail('Malformed or misqualified font size');
  const value=attribute(node,'val',W);
  if(value===undefined||!/^\d+$/.test(value)||!Number.isSafeInteger(Number(value))||Number(value)<=0)fail('Unsupported direct half-point font size');
  return Number(value)/2;
}
function flagValue(node:XmlElement):boolean{
  for(const key of Object.keys(node.attributes))if(node.attributeNamespaces[key]!==XMLNS&&!(key.split(':').at(-1)==='val'&&node.attributeNamespaces[key]===W))fail('Malformed or misqualified Boolean property');
  const value=attribute(node,'val',W);
  if(value===undefined||['1','true','on'].includes(value))return true;
  if(['0','false','off'].includes(value))return false;
  return fail('Unknown Boolean property value');
}
function assertWhitespaceGaps(node:XmlElement,xml:string){
  if(node.selfClosing)return;let cursor=node.openEnd;
  for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,child.start)))fail('Mixed text or lexical barriers are unsupported');cursor=child.end;}
  if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,node.closeStart)))fail('Mixed text or lexical barriers are unsupported');
}
function hasRevision(node:XmlElement):boolean{return node.namespaceURI===W&&node.localName.endsWith('Change')||node.children.some(hasRevision);}
