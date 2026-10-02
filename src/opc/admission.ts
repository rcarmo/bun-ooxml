import {OoxmlError} from '../errors.ts';
import {validateXml} from '../xml/index.ts';
import {decodeXmlBytes} from '../xml/encoding.ts';
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
  validateXml(decodeXmlBytes(payload));
 }
 return parts;
}

/** Validate configuration before obtaining file metadata or bytes.
 * The legacy byte-buffer entry point remains unchanged.
 */
export async function admitPackageFile(path:string,limits:ZipLimits={}):Promise<Map<string,Uint8Array>> {
 validateLimits(limits);
 if(typeof path!=='string'||!path)throw new OoxmlError('package-admission-source-invalid','Expected a nonempty source path');
 const file=Bun.file(path);
 const max=limits.maxArchiveBytes??256*1024*1024;
 if(file.size>max)throw new OoxmlError('zip-archive-too-large','ZIP archive bytes exceed the configured limit');
 return admitPackage(await file.bytes(),limits);
}

function validateLimits(limits:ZipLimits):void {
 for(const key of ['maxEntries','maxEntryBytes','maxTotalBytes','maxArchiveBytes'] as const){
  const value=limits[key];
  if(value!==undefined&&(!Number.isSafeInteger(value)||value<0))throw new OoxmlError('package-admission-limit-invalid',`${key} must be a nonnegative safe integer`);
 }
 const ratio=limits.maxCompressionRatio;
 if(ratio!==undefined&&(!Number.isFinite(ratio)||ratio<=0))throw new OoxmlError('package-admission-limit-invalid','maxCompressionRatio must be finite and positive');
}
