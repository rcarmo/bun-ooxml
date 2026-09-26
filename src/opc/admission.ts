import {OoxmlError} from '../errors.ts';
import {validateXml} from '../xml/index.ts';
import {readZip,type ZipLimits} from './zip.ts';

/** Admit ZIP members and XML syntax without requiring an OPC relationship graph.
 * No filesystem writes, XML entity loading, repair, schema or semantic comparison.
 * Returns detached decompressed payloads; the caller's archive is never modified.
 */
export function admitPackage(bytes:Uint8Array,limits:ZipLimits={}):Map<string,Uint8Array> {
 validateLimits(limits);
 const parts=readZip(bytes,limits);
 for(const [name,payload]of parts){
  if(!/\.(?:xml|rels)$/i.test(name))continue;
  validateXml(decodeAdmissionXml(payload));
 }
 return parts;
}

function validateLimits(limits:ZipLimits):void {
 for(const key of ['maxEntries','maxEntryBytes','maxTotalBytes','maxArchiveBytes'] as const){
  const value=limits[key];
  if(value!==undefined&&(!Number.isSafeInteger(value)||value<0))throw new OoxmlError('package-admission-limit-invalid',`${key} must be a nonnegative safe integer`);
 }
 const ratio=limits.maxCompressionRatio;
 if(ratio!==undefined&&(!Number.isFinite(ratio)||ratio<=0))throw new OoxmlError('package-admission-limit-invalid','maxCompressionRatio must be finite and positive');
}

function decodeAdmissionXml(bytes:Uint8Array):string {
 const encoding=bytes[0]===0xff&&bytes[1]===0xfe?'utf-16le':bytes[0]===0xfe&&bytes[1]===0xff?'utf-16be':'utf-8';
 let text:string;
 try{text=new TextDecoder(encoding,{fatal:true}).decode(bytes);}
 catch{throw new OoxmlError('opc-xml-encoding','Invalid XML member encoding');}
 const declaration=/^<\?xml\s[^?]*\bencoding\s*=\s*(["'])([^"']+)\1/.exec(text)?.[2]?.toLowerCase();
 if(declaration&&declaration!==encoding&&!(declaration==='utf-16'&&encoding!=='utf-8'))throw new OoxmlError('opc-xml-encoding',`XML declaration ${declaration} does not match supported ${encoding} bytes`);
 return text;
}
