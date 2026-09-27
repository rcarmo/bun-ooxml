import assert from 'node:assert/strict';
import {crc32,inflateRawSync} from 'node:zlib';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {OoxmlError} from '../../src/errors.ts';
import {readZip} from '../../src/opc/zip.ts';
import {buildZip} from '../fixtures/zip32.ts';

const encoder=new TextEncoder();
const signatureCRC=0x08074b50;
type State={archive?:Uint8Array;before?:Uint8Array;decoded?:Uint8Array;descriptor?:number;central?:number;geometry?:boolean;error?:unknown;output?:Map<string,Uint8Array>};
const state=(c:Record<string,unknown>)=>c.state as State;
const view=(bytes:Uint8Array)=>new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
const signature=(v:DataView,at:number)=>v.getUint32(at,true);
const find=(v:DataView,length:number,magic:number)=>{for(let at=0;at<=length-4;at++)if(signature(v,at)===magic)return at;throw Error('Missing expected ZIP record');};

export const bindings:StepBinding[]=[
 {pattern:/^a single-disk ZIP32 archive has one DEFLATED data.bin member with declared size 7 and stored payload "payload"$/,run:c=>{
  const s=state(c),payload=encoder.encode('payload');s.archive=buildZip([{name:'data.bin',blob:payload,flags:0x0808,dataDescriptor:true,descriptorSignature:false,crc:signatureCRC}]);s.before=s.archive.slice();
  const v=view(s.archive),central=find(v,s.archive.length,0x02014b50),end=find(v,s.archive.length,0x06054b50);
  assert.equal(v.getUint16(end+4,true),0);assert.equal(v.getUint16(end+6,true),0);assert.equal(v.getUint16(end+10,true),1);
  assert.equal(v.getUint16(central+10,true),8);assert.equal(v.getUint32(central+24,true),7);
 }},
 {pattern:/^its unsigned twelve-byte data descriptor and central directory both declare CRC32 08074B50, equal to the optional descriptor signature value$/,run:c=>{
  const s=state(c);assert(s.archive);const v=view(s.archive),central=find(v,s.archive.length,0x02014b50),local=v.getUint32(central+42,true);
  assert.equal(signature(v,local),0x04034b50);assert.equal(v.getUint16(local+6,true)&8,8);
  const start=local+30+v.getUint16(local+26,true)+v.getUint16(local+28,true),descriptor=start+v.getUint32(central+20,true);
  assert.equal(central-descriptor,12);assert.equal(signature(v,descriptor),signatureCRC);assert.equal(signature(v,central+16),signatureCRC);
  assert.equal(v.getUint32(descriptor+4,true),v.getUint32(central+20,true));assert.equal(v.getUint32(descriptor+8,true),v.getUint32(central+24,true));
  s.central=central;s.descriptor=descriptor;s.decoded=inflateRawSync(s.archive.slice(start,descriptor));s.geometry=true;
 }},
 {pattern:/^independent CRC32 of the decompressed payload differs from 08074B50$/,run:c=>{const s=state(c);assert(s.decoded);assert.deepEqual([...s.decoded],[...encoder.encode('payload')]);assert.notEqual(crc32(s.decoded)>>>0,signatureCRC);}},
 {pattern:/^ZIP admission validates the descriptor shape and then opens the archive with default bounds$/,run:c=>{
  const s=state(c);assert(s.archive);assert(s.geometry);s.before=s.archive.slice();try{s.output=readZip(s.archive);}catch(error){s.error=error;}
 }},
 {pattern:/^the unsigned descriptor is recognised as twelve bytes without borrowing a four-byte signature or central-directory bytes$/,run:c=>{
  const s=state(c);assert(s.archive);assert(s.geometry);assert.equal(s.central!-s.descriptor!,12);
  const v=view(s.archive);assert.equal(signature(v,s.descriptor!),signatureCRC);assert.equal(signature(v,s.central!),0x02014b50);
  assert(s.error instanceof OoxmlError);assert.equal(s.error.code,'zip-crc-mismatch');
 }},
 {pattern:/^complete admission refuses the corrupt payload as a CRC or invalid-package failure, not as an ambiguous descriptor-shape failure$/,run:c=>{
  const s=state(c);assert(s.error instanceof OoxmlError);assert.equal(s.error.code,'zip-crc-mismatch');assert.notEqual(s.error.code,'zip-local-metadata-mismatch');
 }},
 {pattern:/^no package or member payloads are delivered$/,run:c=>{const s=state(c);assert.equal(s.output,undefined);assert(s.error);}},
 {pattern:/^the caller's original archive bytes remain unchanged$/,run:c=>{const s=state(c);assert(s.archive);assert.deepEqual(s.archive,s.before);}},
];
