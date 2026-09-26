import {OoxmlError} from '../errors.ts';
import {applyEdits,attribute,type XmlElement} from '../xml/index.ts';

const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS='http://www.w3.org/2000/xmlns/';
const ORDER=['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike','outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color','spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText','vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath'];
type Edit={start:number;end:number;value:string};
export type DirectRunPatch={bold?:boolean|null;italic?:boolean|null};
const fail=(message:string):never=>{throw new OoxmlError('docx-format-unsupported',message);};
const word=(node:XmlElement,name:string)=>node.namespaceURI===W&&node.localName===name;

/** Preserve the original XML except selected b/i elements or a missing rPr. */
export function formatRunProperties(xml:string,paragraph:XmlElement,patch:DirectRunPatch):{xml:string;changedRuns:number}{
  if(!patch||typeof patch!=='object'||Array.isArray(patch)||Object.keys(patch).some(k=>k!=='bold'&&k!=='italic')||Object.values(patch).some(v=>v!==undefined&&v!==null&&typeof v!=='boolean')){
    throw new OoxmlError('docx-format-argument','Only boolean/null bold and italic overrides are supported');
  }
  const requested=(['bold','italic'] as const).filter(k=>patch[k]!==undefined);
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
        if(node.localName==='b'||node.localName==='i')flagValue(node);
        if(!node.selfClosing&&xml.slice(node.openEnd,node.closeStart).trim())fail('Run properties must not contain text or markup');
      }
    }
    let changed=false;const insertions=new Map<number,string[]>();
    for(const key of requested){
      const name=key==='bold'?'b':'i',value=patch[key]!,node=pr?.children.find(c=>word(c,name));
      const current=node?flagValue(node):undefined;
      if(value===null){if(node){edits.push({start:node.start,end:node.end,value:''});changed=true;}continue;}
      if(node&&current===value)continue;
      const property=`<w:${name} xmlns:w="${W}" w:val="${value?'1':'0'}"/>`;
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
