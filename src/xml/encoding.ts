import {OoxmlError} from '../errors.ts';
/** Strict XML byte decoding shared by admission and conservative comparison. */
export function decodeXmlBytes(bytes:Uint8Array):string {
 const encoding=bytes[0]===0xff&&bytes[1]===0xfe?'utf-16le':bytes[0]===0xfe&&bytes[1]===0xff?'utf-16be':'utf-8';
 let text:string;
 try{text=new TextDecoder(encoding,{fatal:true}).decode(bytes);}
 catch{throw new OoxmlError('opc-xml-encoding','Invalid XML member encoding');}
 const declaration=/^<\?xml\s[^?]*\bencoding\s*=\s*(["'])([^"']+)\1/.exec(text)?.[2]?.toLowerCase();
 if(declaration&&declaration!==encoding&&!(declaration==='utf-16'&&encoding!=='utf-8'))throw new OoxmlError('opc-xml-encoding',`XML declaration ${declaration} does not match supported ${encoding} bytes`);
 return text;
}
