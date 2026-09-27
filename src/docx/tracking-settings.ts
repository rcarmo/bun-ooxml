import {OoxmlError} from '../errors.ts';
import {OpcPackage,addPart,addRelationship,getContentType} from '../opc/index.ts';
import {applyEdits,attribute,parseXml,type XmlElement} from '../xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const XMLNS='http://www.w3.org/2000/xmlns/';
const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings';
const MIME='application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml';
// ECMA-376 Part 1, CT_Settings sequence. Extension children require a separate policy.
const ORDER=('writeProtection,view,zoom,removePersonalInformation,removeDateAndTime,doNotDisplayPageBoundaries,displayBackgroundShape,printPostScriptOverText,printFractionalCharacterWidth,printFormsData,embedTrueTypeFonts,embedSystemFonts,saveSubsetFonts,saveFormsData,mirrorMargins,alignBordersAndEdges,bordersDoNotSurroundHeader,bordersDoNotSurroundFooter,gutterAtTop,hideSpellingErrors,hideGrammaticalErrors,activeWritingStyle,proofState,formsDesign,attachedTemplate,linkStyles,stylePaneFormatFilter,stylePaneSortMethod,documentType,mailMerge,revisionView,trackRevisions,doNotTrackMoves,doNotTrackFormatting,documentProtection,autoFormatOverride,styleLockTheme,styleLockQFSet,defaultTabStop,autoHyphenation,consecutiveHyphenLimit,hyphenationZone,doNotHyphenateCaps,showEnvelope,summaryLength,clickAndTypeStyle,defaultTableStyle,evenAndOddHeaders,bookFoldRevPrinting,bookFoldPrinting,bookFoldPrintingSheets,drawingGridHorizontalSpacing,drawingGridVerticalSpacing,displayHorizontalDrawingGridEvery,displayVerticalDrawingGridEvery,doNotUseMarginsForDrawingGridOrigin,drawingGridHorizontalOrigin,drawingGridVerticalOrigin,doNotShadeFormData,noPunctuationKerning,characterSpacingControl,printTwoOnOne,strictFirstAndLastChars,noLineBreaksAfter,noLineBreaksBefore,savePreviewPicture,doNotValidateAgainstSchema,saveInvalidXml,ignoreMixedContent,alwaysShowPlaceholderText,doNotDemarcateInvalidXml,saveXmlDataOnly,useXSLTWhenSaving,saveThroughXslt,showXMLTags,alwaysMergeEmptyNamespace,updateFields,footnotePr,endnotePr,compat,docVars,rsids,attachedSchema,themeFontLang,clrSchemeMapping,doNotIncludeSubdocsInStats,doNotAutoCompressPictures,forceUpgrade,captions,readModeInkLockDown,smartTagType,doNotEmbedSmartTags,decimalSymbol,listSeparator').split(',');
const REPEATED=new Set(['activeWritingStyle','attachedSchema','smartTagType']);
function fail(message:string):never{throw new OoxmlError('docx-tracking-unsupported',message);}
export type TrackingSettingsReceipt={changed:number;changedParts:string[]};
export function validateTrackAuthor(author:string):void {
 if(typeof author!=='string'||!author.trim()||!author.isWellFormed()||/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/u.test(author))throw new OoxmlError('docx-tracking-argument','Expected a nonblank XML-safe tracking author');
}
function whitespace(node:XmlElement,xml:string):void {
 if(node.selfClosing)return;let cursor=node.openEnd;
 for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,child.start)))fail('Mixed settings content or lexical barriers require review');cursor=child.end;}
 if(!/^[ \t\r\n]*$/.test(xml.slice(cursor,node.closeStart)))fail('Mixed settings content or lexical barriers require review');
}
function onOff(node:XmlElement):boolean {
 for(const key of Object.keys(node.attributes))if(node.attributeNamespaces[key]!==XMLNS&&!(node.attributeNamespaces[key]===W&&key.split(':').at(-1)==='val'))fail('Malformed tracking setting attributes');
 if(node.children.length)fail('Tracking setting must be a leaf');
 const value=attribute(node,'val',W);if(value===undefined||['1','true','on'].includes(value))return true;
 if(['0','false','off'].includes(value))return false;return fail('Invalid tracking on/off value');
}
function scan(pkg:OpcPackage){
 const main=pkg.mainPart(),links=pkg.relationships(main).filter(r=>r.type===REL);
 if(links.length>1||links.some(r=>r.external||!r.resolved||/[?#]/.test(r.target)))fail('Settings relationship must be unique, internal and address a whole part');
 const part=links[0]?.resolved;
 for(const name of pkg.names())if(getContentType(pkg,name)===MIME&&name!==part)fail('Unlinked settings parts require review');
 if(!part)return {enabled:false,part:undefined};
 if(getContentType(pkg,part)!==MIME)fail('Unexpected settings content type');
 // Do not mutate a part owned by another relationship, including other formats.
 let owners=0;
 for(const owner of ['',...pkg.names().filter(n=>n!=='[Content_Types].xml'&&!n.endsWith('.rels'))])for(const r of pkg.relationships(owner))if(!r.external&&r.resolved===part){owners++;if(owner!==main||r.type!==REL)fail('Settings part has another owner');}
 if(owners!==1)fail('Settings part must have exactly one owner');
 const xml=pkg.text(part),doc=parseXml(xml),root=doc.root;
 if(root.namespaceURI!==W||root.localName!=='settings')fail('Expected Transitional Word settings');
 for(const key of Object.keys(root.attributes))if(root.attributeNamespaces[key]!==XMLNS)fail('Decorated settings roots require a separate extension policy');
 whitespace(root,xml);let last=-1;const seen=new Set<string>();
 for(const child of root.children){const rank=ORDER.indexOf(child.localName);if(child.namespaceURI!==W||rank<0||rank<last||seen.has(child.localName)&&!REPEATED.has(child.localName))fail('Unknown, duplicate or out-of-order settings');last=rank;seen.add(child.localName);}
 for(const node of doc.elements)if(['trackRevisions','documentProtection'].includes(node.localName)&&(node.namespaceURI!==W||node.parent!==root))fail('Misplaced or misqualified tracking/protection setting');
 const flag=root.children.find(n=>n.localName==='trackRevisions');if(flag)whitespace(flag,xml);
 return {part,xml,root,flag,enabled:flag?onOff(flag):false,protection:root.children.find(n=>n.localName==='documentProtection')};
}
/** The saved Word setting; it does not wrap Bun edits in revision markup. */
export function readTrackingEnabled(pkg:OpcPackage):boolean{return scan(pkg).enabled;}
export function setTrackingEnabled(pkg:OpcPackage,enabled:boolean):TrackingSettingsReceipt {
 const state=scan(pkg);
 if(state.protection){const p=state.protection;for(const key of Object.keys(p.attributes))if(key.split(':').at(-1)==='enforcement'&&p.attributeNamespaces[key]!==W)fail('Misqualified protection enforcement');const value=attribute(p,'enforcement',W);if(!['0','false','off'].includes(value??''))throw new OoxmlError('docx-tracking-protected','Document protection refuses tracking changes');if(p.children.length)fail('Malformed protection metadata');whitespace(p,state.xml!);}
 if(state.enabled===enabled)return {changed:0,changedParts:[]};
 const flag=`<w:trackRevisions xmlns:w="${W}" w:val="${enabled?'1':'0'}"/>`;
 let part=state.part,next:string;
 if(!part){part='word/settings.xml';if(pkg.names().some(n=>n.toLowerCase()===part!.toLowerCase()))fail('The settings part name is already occupied');next=`<?xml version="1.0" encoding="UTF-8"?><w:settings xmlns:w="${W}">${flag}</w:settings>`;}
 else {
  const {root,xml}=state as typeof state & {root:XmlElement;xml:string};
  if(state.flag)next=applyEdits(xml,[{start:state.flag.start,end:state.flag.end,value:flag}]);
  else if(root.selfClosing)next=applyEdits(xml,[{start:root.start,end:root.end,value:xml.slice(root.start,root.end).replace(/\/[ \t\r\n]*>$/,'>')+flag+`</${root.name}>`}]);
  else {const start=root.children.find(n=>ORDER.indexOf(n.localName)>ORDER.indexOf('trackRevisions'))?.start??root.closeStart;next=applyEdits(xml,[{start,end:start,value:flag}]);}
 }
 const before=new Map(pkg.names().map(n=>[n,pkg.get(n)!]));
 return pkg.transaction(()=>{
  if(state.part){const bytes=pkg.get(part!)!;pkg.set(part!,bytes[0]===239&&bytes[1]===187&&bytes[2]===191?'\ufeff'+next:next);}else{addPart(pkg,part!,next,MIME);addRelationship(pkg,pkg.mainPart(),REL,'settings.xml');}
  if(readTrackingEnabled(pkg)!==enabled)fail('Tracking setting did not survive the edit');
  pkg.toBytes();
  return {changed:1,changedParts:pkg.names().filter(n=>!before.has(n)||!Buffer.from(before.get(n)!).equals(pkg.get(n)!))};
 });
}
