import {OpcPackage} from '../opc/package.ts';
import {parseXml,elements,attribute,escapeText,escapeAttribute,applyEdits,type XmlElement} from '../xml/index.ts';
import {OoxmlError} from '../errors.ts';
import {storyParts,inspectStories} from './story.ts';
import {inspectRevisions,resolveRevisions} from './revisions.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICE='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
function fail(code:string,message:string):never{throw new OoxmlError(code,message);}

/** Author one bounded run-level redline. Queries cannot cross paragraphs or any
 * unreadable run content. Mutation is staged in a private package and committed
 * only after accept/reject copies reproduce their expected story text. Full Word
 * Compare, comments, property/move/table revisions remain separate APIs/gaps.
 */
export function trackedReplace(pkg:OpcPackage,part:string,query:string,replacement:string,options:{author:string;date:string}):{revisionIds:string[];changedParts:string[]}{
 if(typeof query!=='string'||!query||/[\r\n\t]/.test(query)||typeof replacement!=='string'||/[\r\n\t]/.test(replacement))fail('docx-redline-argument','Tracked text must be a nonempty single-paragraph query and plain replacement');
 if(!options||typeof options.author!=='string'||!options.author.trim()||typeof options.date!=='string'||!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d{1,3})?Z$/.test(options.date)||!Number.isFinite(Date.parse(options.date)))fail('docx-redline-argument','Explicit author and UTC ISO date are required');
 const canonical=new Date(options.date).toISOString();
 const normalizedDate=options.date.replace(/(?:\.(\d{1,3}))?Z$/,(_match,fraction:string|undefined)=>'.'+(fraction??'').padEnd(3,'0')+'Z');
 if(canonical!==normalizedDate)fail('docx-redline-argument','Date contains invalid calendar fields');
 escapeText(query);escapeText(replacement);escapeAttribute(options.author);
 const parts=storyParts(pkg);if(!parts.some(p=>p.part===part))fail('docx-redline-story','Target is not a reachable Word story');
 for(const rel of pkg.relationships(pkg.mainPart()))if(!rel.external&&rel.type===OFFICE+'settings'){
  for(const protection of elements(parseXml(pkg.text(rel.resolved!)),'documentProtection',W))if(!['0','false','off'].includes(attribute(protection,'enforcement',W)??''))fail('docx-redline-protected','Protected document refuses tracked replacement');
 }
 const revisionState=inspectRevisions(pkg);
 if(revisionState.revisions.some(r=>r.part===part)||revisionState.unsupported.some(r=>r.part===part))fail('docx-redline-existing-revisions','Resolve target story revisions before redlining');
 const source=pkg.text(part),document=parseXml(source);
 if(document.elements.some(e=>e.namespaceURI===W&&['txbxContent','sdt','sdtContent'].includes(e.localName)||e.namespaceURI==='http://schemas.openxmlformats.org/markup-compatibility/2006'&&e.localName==='AlternateContent'))fail('docx-redline-topology','Text boxes, controls and alternate content require another edit API');
 const candidates:{paragraph:XmlElement;runs:XmlElement[];texts:string[];start:number;fullText:string}[]=[];
 let blind=false;
 for(const paragraph of elements(document,'p',W)){
  const runs=paragraph.children.filter(e=>e.namespaceURI===W&&e.localName==='r');
  const safe=!unownedContent(paragraph,source)&&paragraph.children.every(e=>e.namespaceURI===W&&['r','pPr'].includes(e.localName))&&runs.every(r=>!unownedContent(r,source)&&r.children.every(e=>e.namespaceURI===W&&['rPr','t'].includes(e.localName))&&r.children.filter(e=>e.localName==='t'&&e.namespaceURI===W).every(t=>!t.children.length));
  if(!safe){blind=true;continue;}
  const texts=runs.map(r=>r.children.filter(e=>e.namespaceURI===W&&e.localName==='t').map(e=>e.text).join('')),fullText=texts.join('');
  for(let i=0;(i=fullText.indexOf(query,i))>=0;i+=1)candidates.push({paragraph,runs,texts,start:i,fullText});
 }
 // A drawing/textbox can carry a duplicate invisible target. Refuse incomplete
 // target resolution rather than picking a safe occurrence and hiding blind areas.
 if(blind)fail('docx-redline-topology','Target story contains unsupported paragraph content');
 if(candidates.length!==1)fail('docx-redline-target',`Expected one exact target, found ${candidates.length}`);
 const candidate=candidates[0]!;
 if(query===replacement)return {revisionIds:[],changedParts:[]};
 const ids=new Set(revisionState.revisions.map(r=>r.id));let next=0;
 const nextId=()=>{while(ids.has(String(next)))next++;const id=String(next++);ids.add(id);return id;};
 const delId=nextId(),insId=replacement?nextId():undefined;
 let offset=0;const selected:{run:XmlElement;text:string;left:number;right:number}[]=[];
 for(let i=0;i<candidate.runs.length;i++){const text=candidate.texts[i]!,end=offset+text.length;
  if(text.length>0&&end>candidate.start&&offset<candidate.start+query.length)selected.push({run:candidate.runs[i]!,text,left:Math.max(0,candidate.start-offset),right:Math.min(text.length,candidate.start+query.length-offset)});
  offset=end;
 }
 if(!selected.length)fail('docx-redline-target','Target has no plain text run');
 const first=selected[0]!,last=selected.at(-1)!;
 const covered=candidate.runs.slice(candidate.runs.indexOf(first.run),candidate.runs.indexOf(last.run)+1);
 if(covered.length!==selected.length)fail('docx-redline-topology','A non-text run separates the target fragments');
 const metadata=`xmlns:w="${W}" w:author="${escapeAttribute(options.author)}" w:date="${escapeAttribute(options.date)}"`;
 const deleted=selected.map(s=>renderRun(source,s.run,s.text.slice(s.left,s.right),true)).join('');
 const inserted=insId?`<w:ins ${metadata} w:id="${insId}">${renderRun(source,first.run,replacement,false)}</w:ins>`:'';
 const replacementXml=(first.left?renderRun(source,first.run,first.text.slice(0,first.left),false):'')+`<w:del ${metadata} w:id="${delId}">${deleted}</w:del>`+inserted+(last.right<last.text.length?renderRun(source,last.run,last.text.slice(last.right),false):'');
 for(let index=1;index<selected.length;index++)if(!/^[ \t\r\n]*$/.test(source.slice(selected[index-1]!.run.end,selected[index]!.run.start)))fail('docx-redline-topology','Unowned inter-run content separates the target');
 const nextXml=applyEdits(source,[{start:first.run.start,end:last.run.end,value:replacementXml}]);
 const privateParts=new Map(pkg.names().map(n=>[n,pkg.get(n)!]));const staged=OpcPackage.fromParts(privateParts);staged.set(part,nextXml);
 const beforeText=inspectStories(pkg,{view:'current'}).stories.find(s=>s.part===part)!.paragraphs.map(p=>p.text);
 const expected=beforeText.slice();const paragraphIndex=elements(document,'p',W).indexOf(candidate.paragraph);
 if(paragraphIndex>=expected.length)fail('docx-redline-scope','Story paragraph enumeration cannot safely identify the target');
 expected[paragraphIndex]=candidate.fullText.slice(0,candidate.start)+replacement+candidate.fullText.slice(candidate.start+query.length);
 for(const action of ['accept','reject'] as const){const copy=OpcPackage.fromParts(new Map(staged.names().map(n=>[n,staged.get(n)!])));resolveRevisions(copy,action,{parts:[part]});const actual=inspectStories(copy).stories.find(s=>s.part===part)!.paragraphs.map(p=>p.text);
  if(JSON.stringify(actual)!==JSON.stringify(action==='accept'?expected:beforeText))fail('docx-redline-algebra','Accept/reject verification did not reproduce expected text');
 }
 pkg.transaction(()=>{pkg.set(part,nextXml);pkg.toBytes();});
 return {revisionIds:insId?[delId,insId]:[delId],changedParts:[part]};
}
function unownedContent(node:XmlElement,xml:string):boolean{
 if(node.selfClosing)return false;let cursor=node.openEnd;
 for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,child.start)))return true;cursor=child.end;}
 return !/^[ \t\r\n]*$/.test(xml.slice(cursor,node.closeStart));
}
function renderRun(xml:string,run:XmlElement,text:string,deleted:boolean):string{
 const bindings:Record<string,string>=Object.create(null);const chain:XmlElement[]=[];for(let n:XmlElement|undefined=run;n;n=n.parent)chain.push(n);
 for(const n of chain.reverse())for(const [k,v]of Object.entries(n.attributes))if(k==='xmlns'||k.startsWith('xmlns:'))bindings[k]=v;
 // A private prefix avoids reinterpreting a user's locally rebound lexical w.
 let prefix='rv';for(let i=0;bindings['xmlns:'+prefix]&&bindings['xmlns:'+prefix]!==W;i++)prefix='rv'+i;
 bindings['xmlns:'+prefix]=W;
 const attrs={...bindings,...run.attributes};const properties=run.children.find(e=>e.namespaceURI===W&&e.localName==='rPr');
 const contentName=prefix+':'+(deleted?'delText':'t');
 return `<${prefix}:r${Object.entries(attrs).map(([k,v])=>` ${k}="${escapeAttribute(v)}"`).join('')}>${properties?xml.slice(properties.start,properties.end):''}<${contentName} xml:space="preserve">${escapeText(text)}</${contentName}></${prefix}:r>`;
}
