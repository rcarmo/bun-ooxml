import {crc32} from '../../src/opc/zip.ts';
export const utf8=(value:string)=>new TextEncoder().encode(value);
export function utf16(value:string,bigEndian=false):Uint8Array {
 const out=new Uint8Array(2+value.length*2),view=new DataView(out.buffer);out.set(bigEndian?[0xfe,0xff]:[0xff,0xfe]);
 for(let i=0;i<value.length;i++)view.setUint16(2+i*2,value.charCodeAt(i),!bigEndian);return out;
}
// bzip2 1.0.8 compression of the four UTF-8 bytes <a/>. Offline fixture payload;
// construction and tests never invoke another runtime or compressor.
export const BZIP2_A=Uint8Array.from(Buffer.from('425a6839314159265359996746ea00000099000000800520002000219a68334d173c5dc914e14242659d1ba8','hex'));
/** Ordered ZIP32 builder: permits duplicate/unsafe names so admission sees them. */
export function admissionZip(entries:readonly (readonly [string,Uint8Array])[],method:0|8|12=0):Uint8Array {
 const local:Uint8Array[]=[],central:Uint8Array[]=[];let offset=0;
 for(const [name,payload] of entries){
  const raw=utf8(name),data=method===8?Bun.deflateSync(Uint8Array.from(payload)):method===12?BZIP2_A:payload;
  if(method===12&&Buffer.compare(Buffer.from(payload),Buffer.from(utf8('<a/>')))!==0)throw Error('BZIP2 fixture supports <a/> only');
  const h=new Uint8Array(30),v=new DataView(h.buffer),c=new Uint8Array(46),w=new DataView(c.buffer),crc=crc32(payload);
  v.setUint32(0,0x04034b50,true);v.setUint16(4,method===12?46:20,true);v.setUint16(6,0x800,true);v.setUint16(8,method,true);v.setUint32(14,crc,true);v.setUint32(18,data.length,true);v.setUint32(22,payload.length,true);v.setUint16(26,raw.length,true);
  w.setUint32(0,0x02014b50,true);w.setUint16(4,method===12?46:20,true);w.setUint16(6,method===12?46:20,true);w.setUint16(8,0x800,true);w.setUint16(10,method,true);w.setUint32(16,crc,true);w.setUint32(20,data.length,true);w.setUint32(24,payload.length,true);w.setUint16(28,raw.length,true);w.setUint32(42,offset,true);
  local.push(h,raw,data);central.push(c,raw);offset+=h.length+raw.length+data.length;
 }
 const centralSize=central.reduce((n,b)=>n+b.length,0),end=new Uint8Array(22),e=new DataView(end.buffer);
 e.setUint32(0,0x06054b50,true);e.setUint16(8,entries.length,true);e.setUint16(10,entries.length,true);e.setUint32(12,centralSize,true);e.setUint32(16,offset,true);
 const all=[...local,...central,end],out=new Uint8Array(offset+centralSize+end.length);let i=0;for(const b of all){out.set(b,i);i+=b.length;}return out;
}
