import {expect} from "bun:test";
import {readZip,writeZip,type ZipLimits} from "../../src/opc/zip.ts";
import {OoxmlError} from "../../src/errors.ts";
import type {StepBinding} from "../../scripts/gherkin.ts";
const encoder=new TextEncoder(),decoder=new TextDecoder();
const initial=()=>new Map([["one.xml",encoder.encode("<one>before</one>")],["two.xml",encoder.encode("<two>preserve</two>")]]);
export const bindings:StepBinding[]=[
 {pattern:/^a ZIP64 archive with two native-written XML members$/,run:ctx=>{ctx.bytes=writeZip(initial(),{forceZip64:true});}},
 {pattern:/^Bun reads it and rewrites a changed member with ZIP64 enabled$/,run:ctx=>{const parts=readZip(ctx.bytes as Uint8Array);parts.set("one.xml",encoder.encode("<one>after</one>"));ctx.output=writeZip(parts,{forceZip64:true});}},
 {pattern:/^the changed member reopens with its new payload$/,run:ctx=>{expect(decoder.decode(readZip(ctx.output as Uint8Array).get("one.xml"))).toBe("<one>after</one>");}},
 {pattern:/^the unrelated member retains its exact payload bytes$/,run:ctx=>{expect(readZip(ctx.output as Uint8Array).get("two.xml")).toEqual(initial().get("two.xml")!);}},
 {pattern:/^a ZIP64 archive declaring more entries than its directory can contain$/,run:ctx=>{const b=writeZip(initial(),{forceZip64:true});const v=new DataView(b.buffer,b.byteOffset,b.byteLength);const end=b.length-22-20-56;v.setBigUint64(end+24,100n,true);v.setBigUint64(end+32,100n,true);ctx.bytes=b;}},
 {pattern:/^a ZIP64 locator offset above the safe integer range$/,run:ctx=>{const b=writeZip(initial(),{forceZip64:true});new DataView(b.buffer,b.byteOffset,b.byteLength).setBigUint64(b.length-22-20+8,BigInt(Number.MAX_SAFE_INTEGER)+1n,true);ctx.bytes=b;}},
 {pattern:/^a valid ZIP64 archive with two entries and a one-entry budget$/,run:ctx=>{ctx.bytes=writeZip(initial(),{forceZip64:true});ctx.limits={maxEntries:1};}},
 {pattern:/^Bun attempts bounded ZIP64 admission$/,run:ctx=>{try{readZip(ctx.bytes as Uint8Array,ctx.limits as ZipLimits);ctx.refusal=undefined;}catch(e){ctx.refusal=e;}}},
 {pattern:/^admission refuses with code "([^"]+)"$/,run:(ctx,code)=>{expect(ctx.refusal).toBeInstanceOf(OoxmlError);expect((ctx.refusal as OoxmlError).code).toBe(code);}},
];
