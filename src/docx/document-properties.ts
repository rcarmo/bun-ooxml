import {OoxmlError} from '../errors.ts';
import {parseXml,elements,attribute,applyEdits,type XmlElement} from '../xml/index.ts';
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main',XMLNS='http://www.w3.org/2000/xmlns/';
const ORDER=['headerReference','footerReference','footnotePr','endnotePr','type','pgSz','pgMar','paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote','titlePg','textDirection','bidi','rtlGutter','docGrid','printerSettings'];
export interface DirectDocumentProperties {titlePage:boolean|null;backgroundColor:string|null;}
export type DocumentPropertiesPatch=Partial<DirectDocumentProperties>;
function fail(message:string):never{throw new OoxmlError('docx-document-properties-unsupported',message);}
const word=(node:XmlElement,name:string)=>node.namespaceURI===W&&node.localName===name;
function color(value:unknown):string {if(typeof value!=='string'||!(/^(?:[A-Fa-f0-9]{6}|auto)$/.test(value))||/[\r\n]/.test(value))fail('Background color must be six-digit RGB or auto');return value;}
function normalize(input:DocumentPropertiesPatch):DocumentPropertiesPatch{
 if(!input||typeof input!=='object'||![Object.prototype,null].includes(Object.getPrototypeOf(input)))fail('Expected plain document-property patch');const patch:DocumentPropertiesPatch={};
 for(const key of Reflect.ownKeys(input)){if(key!=='titlePage'&&key!=='backgroundColor')fail('Unknown document property');const d=Object.getOwnPropertyDescriptor(input,key)!;if(!('value'in d))fail('Property accessors are unsupported');const v=d.value;if(v===undefined)continue;if(key==='titlePage'){if(v!==null&&typeof v!=='boolean')fail('Title page must be Boolean or null');patch.titlePage=v;}else patch.backgroundColor=v===null?null:color(v);}
 if(!Object.keys(patch).length)fail('Expected at least one defined property');return patch;
}
function gaps(xml:string,node:XmlElement){if(node.selfClosing)return;let pos=node.openEnd;for(const child of node.children){if(!/^[ \t\r\n]*$/.test(xml.slice(pos,child.start)))fail('Mixed or lexical property content');pos=child.end;}if(!/^[ \t\r\n]*$/.test(xml.slice(pos,node.closeStart)))fail('Mixed or lexical property content');}
function attrs(node:XmlElement,allowed:string[]){for(const key of Object.keys(node.attributes))if(node.attributeNamespaces[key]!==XMLNS&&!(node.attributeNamespaces[key]===W&&allowed.includes(key.split(':').at(-1)!)))fail('Unsupported or misqualified property attribute');}
function metadata(xml:string,node:XmlElement){if(node.namespaceURI!==W||node.localName.endsWith('Change')||['ins','del'].includes(node.localName))fail('Unsupported or revised section metadata');gaps(xml,node);for(const child of node.children)metadata(xml,child);}
function inspect(xml:string){
 const tree=parseXml(xml),root=tree.root;if(!word(root,'document'))fail('Expected Word document root');gaps(xml,root);
 const bodies=root.children.filter(n=>word(n,'body')),backgrounds=root.children.filter(n=>word(n,'background'));
 if(bodies.length!==1||backgrounds.length>1||root.children.some(n=>!word(n,'body')&&!word(n,'background'))||root.children.at(-1)!==bodies[0])fail('Expected optional background before a unique body');
 const body=bodies[0]!,sections=elements(tree,'sectPr',W);if(sections.length!==1||sections[0]!.parent!==body||body.children.at(-1)!==sections[0])fail('Expected exactly one final section in the document');
 const section=sections[0]!;gaps(xml,section);let last=-1;const seen=new Set<string>();
 for(const child of section.children){const rank=ORDER.indexOf(child.localName);if(child.namespaceURI!==W||rank<0||rank<last||(seen.has(child.localName)&&!['headerReference','footerReference'].includes(child.localName)))fail('Unknown, duplicate or misordered section metadata');last=rank;seen.add(child.localName);metadata(xml,child);}
 const title=section.children.find(n=>word(n,'titlePg')),background=backgrounds[0];let titlePage:boolean|null=null,backgroundColor:string|null=null;
 if(title){attrs(title,['val']);if(title.children.length)fail('Title-page marker must be empty');const raw=attribute(title,'val',W);if(raw===undefined||['1','true','on'].includes(raw))titlePage=true;else if(['0','false','off'].includes(raw))titlePage=false;else fail('Invalid title-page Boolean');}
 if(background){attrs(background,['color']);gaps(xml,background);if(background.children.length)fail('Drawing backgrounds are unsupported');const raw=attribute(background,'color',W);backgroundColor=raw===undefined?null:color(raw);}
 return {body,section,title,background,values:{titlePage,backgroundColor}};
}
export function readDocumentProperties(xml:string):DirectDocumentProperties{return inspect(xml).values;}
export function editDocumentProperties(xml:string,input:DocumentPropertiesPatch):{xml:string;changed:number}{
 const patch=normalize(input),state=inspect(xml),edits:Array<{start:number;end:number;value:string}>=[];let changed=0;
 if(patch.titlePage!==undefined){const value=patch.titlePage,node=state.title;if(value===null?!!node:value!==state.values.titlePage){changed++;const fragment=value===null?'':`<w:titlePg xmlns:w="${W}" w:val="${value?'1':'0'}"/>`;
  if(node)edits.push({start:node.start,end:node.end,value:fragment});else if(state.section.selfClosing){const n=state.section;edits.push({start:n.start,end:n.end,value:xml.slice(n.start,n.end).replace(/\/>$/,'>')+fragment+`</${n.name}>`});}
  else{const at=state.section.children.find(n=>ORDER.indexOf(n.localName)>ORDER.indexOf('titlePg'))?.start??state.section.closeStart;edits.push({start:at,end:at,value:fragment});}
 }}
 if(patch.backgroundColor!==undefined){const value=patch.backgroundColor,node=state.background;if(value===null?!!node:value!==state.values.backgroundColor){changed++;const fragment=value===null?'':`<w:background xmlns:w="${W}" w:color="${value}"/>`;if(node)edits.push({start:node.start,end:node.end,value:fragment});else edits.push({start:state.body.start,end:state.body.start,value:fragment});}}
 return {xml:edits.length?applyEdits(xml,edits):xml,changed};
}
