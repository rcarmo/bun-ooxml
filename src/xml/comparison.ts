import {OoxmlError} from '../errors.ts';
import {decodeXmlBytes} from './encoding.ts';
import {inspectXmlEvents, type XmlElement} from './index.ts';
const XMLNS='http://www.w3.org/2000/xmlns/';
const REL='http://schemas.openxmlformats.org/package/2006/relationships';
const MC='http://schemas.openxmlformats.org/markup-compatibility/2006';
const XSI='http://www.w3.org/2001/XMLSchema-instance';
type Attribute=[string,string,string];
type Token={kind:'text'|'comment'|'pi'|'declaration';value:string;target?:string}|Element;
type Element={kind:'element';uri:string;local:string;attributes:Attribute[];bindings?:Array<[string,string]>;children:Token[]};

/** Conservative preservation comparison, not canonical XML or signature validation.
 * Unsafe/unsupported inputs return false, including identical malformed bytes.
 */
export function xmlEquivalent(left:Uint8Array,right:Uint8Array):boolean {
 try{return JSON.stringify(model(left))===JSON.stringify(model(right));}
 catch(error){if(error instanceof OoxmlError)return false;throw error;}
}
function model(bytes:Uint8Array):Token[]{
 const source=decodeXmlBytes(bytes);
 // XML 1.1 character rules are not implemented by this parser.
 const version=/^<\?xml\s+version\s*=\s*(["'])([^"']+)\1/.exec(source)?.[2];
 if(version&&version!=='1.0')throw new OoxmlError('xml-comparison-version','Only XML 1.0 comparison is supported');
 const declaration=/^<\?xml\s[^?]*\bstandalone\s*=\s*(["'])(yes|no)\1/.exec(source)?.[2];
 const document:Token[]=declaration?[{kind:'declaration',value:declaration}]:[],stack:Element[]=[];
 const scopes=new WeakMap<Element,ReadonlyMap<string,string>>();
 const append=(node:Token)=>{(stack.at(-1)?.children??document).push(node);};
 inspectXmlEvents(source,{
  start(node:XmlElement,namespaces:ReadonlyMap<string,string>){
   const attributes:Attribute[]=[];let sensitive=false;
   for(const [name,value]of Object.entries(node.attributes)){
    const uri=node.attributeNamespaces[name]!;if(uri===XMLNS)continue;
    const local=name.includes(':')?name.slice(name.indexOf(':')+1):name;
    attributes.push([uri,local,value]);
    const prefixList=uri===MC&&(local==='Ignorable'||local==='MustUnderstand')||node.namespaceURI===MC&&node.localName==='Choice'&&uri===''&&local==='Requires';
    if(value.includes(':')||prefixList||uri===XSI&&local==='type'){
     sensitive=true;
     // Unknown QName-like values are conservatively incomparable, not resolved by guessing.
     if(value.includes(':'))for(const word of value.split(/\s+/)){const prefix=word.split(':')[0]!;if(word.includes(':')&&!namespaces.has(prefix))throw new OoxmlError('xml-comparison-namespace','Unknown prefix-valued attribute binding');}
     if(prefixList)for(const prefix of value.split(/\s+/).filter(Boolean))if(!namespaces.has(prefix))throw new OoxmlError('xml-comparison-namespace','Unknown compatibility prefix');
    }
   }
   attributes.sort((a,b)=>{const x=JSON.stringify(a),y=JSON.stringify(b);return x<y?-1:x>y?1:0;});
   const element:Element={kind:'element',uri:node.namespaceURI,local:node.localName,attributes,children:[]};
   if(sensitive)element.bindings=[...namespaces.entries()].sort(([a],[b])=>a<b?-1:a>b?1:0);
   scopes.set(element,namespaces);append(element);stack.push(element);
  },
  end(){
   const element=stack.pop()!;
   // QName text may be split across CDATA or entity events. Decide after merging.
   if(element.children.some(c=>c.kind==='text'&&c.value.includes(':')))element.bindings=[...scopes.get(element)!.entries()].sort(([a],[b])=>a<b?-1:a>b?1:0);
   if(stack.length===0&&element.uri===REL&&element.local==='Relationships'&&element.children.every(c=>c.kind==='element'&&c.uri===REL&&c.local==='Relationship'&&!c.children.length)){
    const ids=new Set<string>();
    for(const child of element.children as Element[]){const id=child.attributes.find(([uri,local])=>uri===''&&local==='Id')?.[2];if(!id||ids.has(id))throw new OoxmlError('xml-comparison-relationships','Ambiguous relationship collection');ids.add(id);}
    element.children.sort((a,b)=>{const x=JSON.stringify(a),y=JSON.stringify(b);return x<y?-1:x>y?1:0;});
   }
  },
  text(value){const list=stack.at(-1)!.children,last=list.at(-1);if(last?.kind==='text')last.value+=value;else append({kind:'text',value});},
  comment(value){append({kind:'comment',value});},
  instruction(target,value){append({kind:'pi',target,value});},
 });
 return document;
}
