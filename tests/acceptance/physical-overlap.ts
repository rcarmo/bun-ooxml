import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {crc32} from 'node:zlib';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {fixturePath} from '../../scripts/fixture-inputs.ts';
import {admitPackage} from '../../src/opc/admission.ts';
import {OoxmlError} from '../../src/errors.ts';

const fixture='fixture-9286fc07c3f8698f9637cf9b7a0c60461d1f304753ba69b39f655c8c348ea027';
type Member={name:string;size:number;crc:number;local:number;start:number;end:number};
type State={source?:Uint8Array;original?:Uint8Array;members?:Member[];directory?:number;result?:Map<string,Uint8Array>;output?:Uint8Array;error?:unknown};
const state=(c:Record<string,unknown>)=>c.state as State;
const required=<T>(value:T|undefined):T=>{assert(value!==undefined);return value;};
function geometry(bytes:Uint8Array){
 const view=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength),u16=(p:number)=>view.getUint16(p,true),u32=(p:number)=>view.getUint32(p,true),decode=new TextDecoder();
 const eocd=bytes.length-22;assert.equal(u32(eocd),0x06054b50);assert.equal(u16(eocd+4),0);assert.equal(u16(eocd+6),0);assert.equal(u16(eocd+8),3);assert.equal(u16(eocd+10),3);assert.equal(u16(eocd+20),0);
 const directory=u32(eocd+16);assert.equal(directory+u32(eocd+12),eocd);
 const members:Member[]=[];let cursor=directory;
 for(let index=0;index<3;index++){
  assert.equal(u32(cursor),0x02014b50);assert.equal(u16(cursor+8),0);assert.equal(u16(cursor+10),0);
  const size=u32(cursor+20),uncompressed=u32(cursor+24),crc=u32(cursor+16),nameLength=u16(cursor+28),extraLength=u16(cursor+30),commentLength=u16(cursor+32),local=u32(cursor+42),name=decode.decode(bytes.subarray(cursor+46,cursor+46+nameLength));
  assert.equal(size,uncompressed);assert.equal(u32(local),0x04034b50);assert.equal(u16(local+6),0);assert.equal(u16(local+8),0);
  assert.equal(decode.decode(bytes.subarray(local+30,local+30+u16(local+26))),name);
  const start=local+30+u16(local+26)+u16(local+28),end=start+size;
  assert(end<=directory);assert.equal(u32(local+14),crc);assert.equal(u32(local+18),size);assert.equal(u32(local+22),size);
  members.push({name,size,crc,local,start,end});cursor+=46+nameLength+extraLength+commentLength;
 }
 assert.equal(cursor,eocd);return {members,directory};
}
export const bindings:StepBinding[]=[
 {pattern:/^fixture fixture-9286fc07c3f8698f9637cf9b7a0c60461d1f304753ba69b39f655c8c348ea027 has exactly three distinct STORED members \[Content_Types\]\.xml, outer\.bin and inner\.bin$/,run:c=>{
  const s=state(c);s.source=Uint8Array.from(readFileSync(fixturePath(fixture)));s.original=s.source.slice();assert.equal(s.source.length,483);
  const g=geometry(s.source);s.members=g.members;s.directory=g.directory;
  assert.deepEqual(g.members.map(m=>m.name),['[Content_Types].xml','outer.bin','inner.bin']);assert.equal(new Set(g.members.map(m=>m.name)).size,3);
 }},
 {pattern:/^an independent ZIP reader opens all three members with declared lengths 149, 49 and 10 bytes and matching CRC32 d694f44a, 32c80458 and 4daa6380$/,run:c=>{
  const s=state(c),source=required(s.source),members=required(s.members);
  assert.deepEqual(members.map(m=>m.size),[149,49,10]);assert.deepEqual(members.map(m=>m.crc.toString(16).padStart(8,'0')),['d694f44a','32c80458','4daa6380']);
  for(const m of members){const payload=source.subarray(m.start,m.end);assert.equal(payload.length,m.size);assert.equal(crc32(payload)>>>0,m.crc);}
 }},
 {pattern:/^inner\.bin's complete local header and payload lie within outer\.bin's physical payload range in this ZIP32 single-disk archive$/,run:c=>{
  const [types,outer,inner]=required(state(c).members);assert(types&&outer&&inner);
  assert.equal(types.end,outer.local);assert.equal(inner.local,outer.start);assert(inner.local<inner.start&&inner.start<inner.end&&inner.end<=outer.end);
  assert.equal(outer.end,required(state(c).directory));
 }},
 {pattern:/^bounded package admission checks the unchanged archive with max entries 4 and max total bytes 4096$/,run:c=>{
  const s=state(c);assert.deepEqual(s.source,s.original);assert(s.members&&s.directory!==undefined);
  try{s.result=admitPackage(required(s.source),{maxEntries:4,maxTotalBytes:4096});}catch(e){s.error=e;}
 }},
 {pattern:/^admission refuses overlapping physical member extents as an invalid package, not a name, CRC or resource refusal$/,run:c=>{
  const s=state(c);assert(s.error instanceof OoxmlError);assert.equal(s.error.code,'zip-size-mismatch');
  assert.match(s.error.message,/outer\.bin.*declared size overruns its local record/);
  assert(!/name|crc|resource|limit|entries|ratio/i.test(s.error.code));
 }},
 {pattern:/^no package session or output archive is delivered$/,run:c=>{const s=state(c);assert.equal(s.result,undefined);assert.equal(s.output,undefined);}},
 {pattern:/^the caller's source buffer remains byte-identical to the sealed fixture$/,run:c=>{
  const s=state(c);assert.deepEqual(s.source,s.original);assert.deepEqual(Uint8Array.from(readFileSync(fixturePath(fixture))),s.original);
 }},
];
