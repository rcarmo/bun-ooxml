import {OoxmlError} from '../errors.ts';
import {XmlSnapshot,type XmlRemovalTarget} from './removal.ts';
const MAX_BYTES=32*1024*1024;
type Encoding='utf-8'|'utf-16le'|'utf-16be';
const typedArray=Object.getPrototypeOf(Uint8Array.prototype);
const byteLength=Object.getOwnPropertyDescriptor(typedArray,'byteLength')!.get!;
const buffer=Object.getOwnPropertyDescriptor(typedArray,'buffer')!.get!;
function fail(code:string,message:string):never{throw new OoxmlError(code,message);}

/** Owned XML bytes with reusable string-offset handles and encoding-preserving removal. */
export class XmlByteSnapshot {
 readonly elements:readonly XmlRemovalTarget[];
 readonly #bytes:Uint8Array;
 readonly #snapshot:XmlSnapshot;
 readonly #text:string;
 readonly #encoding:Encoding;
 readonly #bom:number;
 private constructor(input:Uint8Array){
  if(!(input instanceof Uint8Array))fail('XML_BYTE_SOURCE','Expected a Uint8Array XML input');
  let length:number,storage:ArrayBufferLike;
  try{length=byteLength.call(input);storage=buffer.call(input);}catch{fail('XML_BYTE_SOURCE','Expected a native byte view');}
  if(length>MAX_BYTES)fail('XML_BYTE_LIMIT','XML byte input exceeds 32 MiB');
  if(typeof SharedArrayBuffer!=='undefined'&&storage instanceof SharedArrayBuffer)fail('XML_BYTE_SOURCE','Shared mutable buffers are not supported');
  // Typed-array copy uses the internal view, not slice, species or its iterator.
  try{this.#bytes=new Uint8Array(input);}catch{fail('XML_BYTE_SOURCE','Detached or inaccessible XML bytes');}
  const bytes=this.#bytes;
  this.#encoding=bytes[0]===255&&bytes[1]===254?'utf-16le':bytes[0]===254&&bytes[1]===255?'utf-16be':'utf-8';
  this.#bom=this.#encoding==='utf-8'?(bytes[0]===239&&bytes[1]===187&&bytes[2]===191?3:0):2;
  if(this.#encoding!=='utf-8'&&(bytes.length-this.#bom)%2!==0)fail('XML_BYTE_ENCODING','Odd UTF-16 byte length');
  try{this.#text=new TextDecoder(this.#encoding,{fatal:true,ignoreBOM:true}).decode(bytes.subarray(this.#bom));}
  catch{fail('XML_BYTE_ENCODING','Malformed XML byte encoding');}
  if(this.#text.startsWith('\ufeff'))fail('XML_BYTE_ENCODING','Duplicate leading BOM is not supported');
  // The scanner validates declaration syntax; this check enforces the byte codec.
  this.#snapshot=XmlSnapshot.parse(this.#text);
  const declared=/^<\?xml[ \t\r\n]+[^?]*?\bencoding[ \t\r\n]*=[ \t\r\n]*(["'])([^"']+)\1/.exec(this.#text)?.[2]?.toLowerCase();
  if(declared&&declared!==this.#encoding&&!(declared==='utf-16'&&this.#encoding!=='utf-8'))fail('XML_BYTE_ENCODING','XML encoding declaration is unsupported or conflicts with byte encoding');
  this.elements=this.#snapshot.elements;
  Object.freeze(this);
 }
 static parse(input:Uint8Array):XmlByteSnapshot{return new XmlByteSnapshot(input);}
 /** New detached output; empty removal preserves the original byte sequence. */
 remove(targets:readonly XmlRemovalTarget[]):Uint8Array{
  const text=this.#snapshot.remove(targets);
  if(text===this.#text)return this.#bytes.slice();
  if(this.#encoding==='utf-8'){
   const payload=new TextEncoder().encode(text),out=new Uint8Array(this.#bom+payload.length);
   out.set(this.#bytes.subarray(0,this.#bom));out.set(payload,this.#bom);return out;
  }
  const out=new Uint8Array(2+text.length*2);out.set(this.#bytes.subarray(0,2));const view=new DataView(out.buffer);
  for(let i=0;i<text.length;i++)view.setUint16(2+i*2,text.charCodeAt(i),this.#encoding==='utf-16le');return out;
 }
}
