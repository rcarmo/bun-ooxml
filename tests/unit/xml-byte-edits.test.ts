import {test,expect} from 'bun:test';
import {XmlByteSnapshot,XmlSnapshot,type XmlContent,type XmlStructurePatch,type XmlAttributePatch} from '../../src/index.ts';
import {parseXml,elements,attribute} from '../../src/xml/index.ts';
type Codec='utf8'|'bom8'|'le'|'be';
const codecs:Codec[]=['utf8','bom8','le','be'];
function encode(text:string,codec:Codec):Uint8Array{
 if(codec==='utf8')return new TextEncoder().encode(text);
 if(codec==='bom8'){const p=encode(text,'utf8'),b=new Uint8Array(p.length+3);b.set([239,187,191]);b.set(p,3);return b;}
 const b=new Uint8Array(2+text.length*2);b.set(codec==='le'?[255,254]:[254,255]);const v=new DataView(b.buffer);for(let i=0;i<text.length;i++)v.setUint16(2+i*2,text.charCodeAt(i),codec==='le');return b;
}
const declaration=(c:Codec)=>`<?xml version="1.0" encoding="${c==='utf8'||c==='bom8'?'UTF-8':'UTF-16'}"?>\r\n`;
const decode=(b:Uint8Array,c:Codec)=>new TextDecoder(c==='le'?'utf-16le':c==='be'?'utf-16be':'utf-8',{fatal:true}).decode(b);
const name=(localName:string,namespaceURI='')=>({localName,namespaceURI});

test('byte attribute batches preserve each codec BOM declaration alias quotes and untouched fragments',()=>{
 for(const c of codecs){const source=declaration(c)+`<r xmlns:q='urn:q' xmlns:z='urn:q'>雪😀<q:a q:id = 'old' keep="&#x41;"/><tail/></r>`,input=encode(source,c),before=input.slice(),s=XmlByteSnapshot.parse(input),target=s.elements[1]!;
  const out=s.setAttributes([{target,name:'z:id',value:'A&B雪😀'},{target,name:'fresh',value:'<new>'}]);
  expect(out).toEqual(encode(source.replace("q:id = 'old'","q:id = 'A&amp;B雪😀'").replace('keep="&#x41;"/','keep="&#x41;" fresh="&lt;new&gt;"/'),c));
  const tree=parseXml(decode(out,c)),node=elements(tree,'a','urn:q')[0]!;expect(attribute(node,'id','urn:q')).toBe('A&B雪😀');expect(node.attributes.fresh).toBe('<new>');
  expect(input).toEqual(before);expect(s.remove([])).toEqual(before);expect(XmlByteSnapshot.parse(out).remove([])).toEqual(out);
 }
});

test('byte append expands self-closing nodes and escapes authored text in surviving namespace scope',()=>{
 for(const c of codecs){const source=declaration(c)+`<r xmlns='urn:r' xmlns:q='urn:q'><slot/><!--keep--><tail a='&#65;'/></r>`,s=XmlByteSnapshot.parse(encode(source,c));
  const out=s.appendChildren([{target:s.elements[1]!,children:[{name:name('child','urn:q'),attributes:[{name:name('id','urn:q'),value:'雪&'}],children:[' A&B😀 ']},{name:name('plain')}]}]);
  expect(out).toEqual(encode(source.replace('<slot/>','<slot><q:child q:id="雪&amp;"> A&amp;B😀 </q:child><plain xmlns=""/></slot>'),c));
  const tree=parseXml(decode(out,c));expect(elements(tree,'child','urn:q')[0]!.text).toBe(' A&B😀 ');expect(elements(tree,'plain','')).toHaveLength(1);expect(s.remove([])).toEqual(encode(source,c));
 }
});

test('byte replacement drops target-local bindings but retains the parent scope and other lexical bytes',()=>{
 for(const c of codecs){const source=declaration(c)+`<r xmlns:q='urn:parent'><q:old xmlns:q="urn:removed"/><keep a='&amp;'/></r>`,s=XmlByteSnapshot.parse(encode(source,c));
  const out=s.replaceElements([{target:s.elements[1]!,children:[{name:name('new','urn:removed'),children:['雪😀']},{name:name('sibling','urn:parent')}]}]);
  expect(out).toEqual(encode(source.replace('<q:old xmlns:q="urn:removed"/>','<n1:new xmlns:n1="urn:removed">雪😀</n1:new><q:sibling/>'),c));
  const parsed=XmlByteSnapshot.parse(out);expect(parsed.elements.map(n=>[n.localName,n.namespaceURI])).toEqual([['r',''],['new','urn:removed'],['sibling','urn:parent'],['keep','']]);expect(s.remove([])).toEqual(encode(source,c));
 }
});

test('empty and semantic no-op byte edits return exact detached buffers while validating their targets',()=>{
 for(const c of codecs){const source=declaration(c)+`<r a = 'A&#x26;B'><leaf/></r>`,before=encode(source,c),s=XmlByteSnapshot.parse(before);
  const outputs=[s.setAttributes([]),s.setAttributes([{target:s.elements[0]!,name:'a',value:'A&B'}]),s.appendChildren([]),s.appendChildren([{target:s.elements[1]!,children:[]}]),s.replaceElements([]),s.replaceElements([{target:s.elements[1]!,children:[{name:name('leaf')}]}])];
  for(const out of outputs){expect(out).toEqual(before);expect(out.buffer).not.toBe(before.buffer);out.fill(0);expect(s.remove([])).toEqual(before);}
  expect(()=>s.appendChildren([{target:{...s.elements[1]!},children:[]}])).toThrow(expect.objectContaining({code:'XML_STRUCTURE_TARGET'}));expect(()=>s.replaceElements([{target:s.elements[0]!,children:[]}])).toThrow(expect.objectContaining({code:'XML_STRUCTURE_ROOT'}));
 }
});

test('byte edits own caller bytes and changed outputs and reuse original handles across operation kinds',()=>{
 for(const c of codecs){const source=declaration(c)+'<r><a/><b/></r>',input=encode(source,c),before=input.slice(),s=XmlByteSnapshot.parse(input),a=s.elements[1]!,b=s.elements[2]!;input.fill(0);structuredClone(input,{transfer:[input.buffer]});
  const ops=[()=>s.setAttributes([{target:a,name:'id',value:'雪'}]),()=>s.appendChildren([{target:a,children:['😀']}]),()=>s.replaceElements([{target:b,children:[{name:name('new')}]}])],expected=[source.replace('<a/>','<a id="雪"/>'),source.replace('<a/>','<a>😀</a>'),source.replace('<b/>','<new/>')];
  for(const [i,op]of ops.entries()){const out=op();expect(out).toEqual(encode(expected[i]!,c));out.fill(0);expect(op()).toEqual(encode(expected[i]!,c));expect(s.remove([])).toEqual(before);}
  expect(s.remove([a])).toEqual(encode(source.replace('<a/>',''),c));
 }
});

test('foreign copied string and byte targets refuse every new operation without invalidating owned handles',()=>{
 const input=encode('<r><a/></r>','le'),before=input.slice(),s=XmlByteSnapshot.parse(input),other=XmlByteSnapshot.parse(input),string=XmlSnapshot.parse('<r><a/></r>');
 for(const target of [other.elements[1]!,string.elements[1]!,{...s.elements[1]!}]){
  expect(()=>s.setAttributes([{target,name:'a',value:'b'}])).toThrow(expect.objectContaining({code:'XML_ATTRIBUTE_TARGET'}));
  for(const op of ['appendChildren','replaceElements'] as const)expect(()=>s[op]([{target,children:[]}])).toThrow(expect.objectContaining({code:'XML_STRUCTURE_TARGET'}));
  expect(s.remove([])).toEqual(before);expect(input).toEqual(before);
 }
 expect(s.setAttributes([{target:s.elements[1]!,name:'a',value:'b'}])).toEqual(encode('<r><a a="b"/></r>','le'));
});

test('byte batches preserve attribute and structure atomic refusals for late invalid data and overlapping targets',()=>{
 const source='<r xmlns:q="urn:q" xmlns:z="urn:q"><a><b/></a><c/></r>',input=encode(source,'be'),s=XmlByteSnapshot.parse(input),a=s.elements[1]!,b=s.elements[2]!,c=s.elements[3]!;
 for(const patch of [{name:'xmlns:p',value:'urn:p'},{name:'unbound:a',value:'x'},{name:'bad',value:'\u0000'},{name:'bad',value:'\ud800'}]){expect(()=>s.setAttributes([{target:a,name:'ok',value:'first'},{target:c,...patch}])).toThrow();expect(s.remove([])).toEqual(input);}
 expect(()=>s.setAttributes([{target:a,name:'q:id',value:'1'},{target:a,name:'z:id',value:'2'}])).toThrow(expect.objectContaining({code:'XML_ATTRIBUTE_DUPLICATE'}));
 const cyclic={name:name('cycle'),children:[] as XmlContent[]};cyclic.children.push(cyclic);
 for(const op of ['appendChildren','replaceElements'] as const){
  for(const children of [[cyclic],['\ud800'],[{name:name('a:b')}],[{name:name('x'),attributes:[{name:name('id'),value:'1'},{name:name('id'),value:'2'}]}]]){expect(()=>s[op]([{target:a,children:['valid first']},{target:c,children}])).toThrow();expect(s.remove([])).toEqual(input);}
  for(const targets of [[a,a],[a,b]])expect(()=>s[op](targets.map(target=>({target,children:[]})))).toThrow(expect.objectContaining({code:'XML_STRUCTURE_OVERLAP'}));
 }
 const unsafe=XmlByteSnapshot.parse(encode('<r>]]<a/>></r>','bom8'));expect(()=>unsafe.replaceElements([{target:unsafe.elements[1]!,children:[]}])).toThrow(expect.objectContaining({code:'XML_STRUCTURE_UNSAFE'}));expect(unsafe.remove([])).toEqual(encode('<r>]]<a/>></r>','bom8'));expect(input).toEqual(encode(source,'be'));
});

test('byte patches keep complete output and nesting limits enforced before encoding',()=>{
 const source='<r><a/> '+' '.repeat(8*1024*1024-12)+'</r>',input=encode(source,'le'),s=XmlByteSnapshot.parse(input),target=s.elements[1]!;
 expect(()=>s.setAttributes([{target,name:'id',value:'too long'}])).toThrow(expect.objectContaining({code:'XML_ATTRIBUTE_LIMIT'}));
 for(const op of ['appendChildren','replaceElements'] as const)expect(()=>s[op]([{target,children:['too long for output']}])).toThrow(expect.objectContaining({code:'XML_STRUCTURE_LIMIT'}));
 expect(s.remove([])).toEqual(input);
 const deep=XmlByteSnapshot.parse(encode('<r>'.repeat(256)+'</r>'.repeat(256),'utf8'));expect(()=>deep.appendChildren([{target:deep.elements[255]!,children:[{name:name('child')}]}])).toThrow(expect.objectContaining({code:'XML_STRUCTURE_LIMIT'}));
});

test('byte edits ignore overridden patch iterators and honour Buffer subview ownership',()=>{
 const backing=Buffer.from('XX<r><a/></r>YY'),view=backing.subarray(2,backing.length-2),s=XmlByteSnapshot.parse(view);backing.fill(0);
 const attrs:XmlAttributePatch[]=[{target:s.elements[1]!,name:'x',value:'y'}];attrs[Symbol.iterator]=()=>{throw Error('iterator');};expect(s.setAttributes(attrs)).toEqual(encode('<r><a x="y"/></r>','utf8'));
 for(const op of ['appendChildren','replaceElements'] as const){const children:XmlContent[]=['ok'];children[Symbol.iterator]=()=>{throw Error('child iterator');};const patches:XmlStructurePatch[]=[{target:s.elements[1]!,children}];patches[Symbol.iterator]=()=>{throw Error('patch iterator');};expect(s[op](patches)).toEqual(encode(op==='appendChildren'?'<r><a>ok</a></r>':'<r>ok</r>','utf8'));}
 expect(s.remove([])).toEqual(encode('<r><a/></r>','utf8'));
});

test('encoded edits can be reparsed for a sequential pipeline without accepting old handles',()=>{
 for(const c of codecs){const s=XmlByteSnapshot.parse(encode(declaration(c)+'<r><a/></r>',c)),one=XmlByteSnapshot.parse(s.setAttributes([{target:s.elements[1]!,name:'id',value:'1'}])),two=XmlByteSnapshot.parse(one.appendChildren([{target:one.elements[1]!,children:[{name:name('b'),children:['雪😀']}]}]));
  expect(()=>two.setAttributes([{target:s.elements[1]!,name:'id',value:'2'}])).toThrow(expect.objectContaining({code:'XML_ATTRIBUTE_TARGET'}));
  const out=two.replaceElements([{target:two.elements[2]!,children:['new']}]);expect(out).toEqual(encode(declaration(c)+'<r><a id="1">new</a></r>',c));expect(s.remove([])).toEqual(encode(declaration(c)+'<r><a/></r>',c));
 }
});
