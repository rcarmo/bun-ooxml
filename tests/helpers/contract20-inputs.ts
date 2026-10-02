import assert from 'node:assert/strict';
import {readZip,writeZip} from '../../src/opc/zip.ts';
import pin from '../../docs/behaviors/contract20-candidate.json';
const encoder=new TextEncoder(),decoder=new TextDecoder();
export const contract20Root=pin.root;
export async function contract20Recipe(id:string):Promise<Uint8Array>{
 const recipes=await Bun.file(contract20Root+'/ledgers/contract20-recipes.json').json();
 const x=recipes.xlsx.find((r:any)=>r.id===id);if(x)return writeZip(new Map(Object.entries(x.members).map(([p,t])=>[p,encoder.encode(t as string)])));
 const p=recipes.pptx.find((r:any)=>r.id===id);assert(p,'known recipe');const manifest=await Bun.file(contract20Root+'/manifest.json').json(),f=manifest.files.find((f:any)=>f.id===p.baseFixtureId);assert(f);
 const bytes=await Bun.file(contract20Root+'/'+f.path).bytes();assert.equal(bytes.length,f.bytes);assert.equal(new Bun.CryptoHasher('sha256').update(bytes).digest('hex'),f.sha256);
 const parts=readZip(bytes);
 for(const op of p.operations){if(op.kind==='add-literal-member'){assert(!parts.has(op.part));parts.set(op.part,encoder.encode(op.value));continue;}
  const text=decoder.decode(parts.get(op.part));let next:string;
  if(op.kind==='replace-literal-once'){assert.equal(text.split(op.before).length,2);next=text.replace(op.before,op.after);}else{assert.equal(op.kind,'reorder-sldId-elements');const matches=[...text.matchAll(/<p:sldId\b[^>]*\/>/g)];assert.equal(matches.length,op.originalCount);const first=matches[0]!,last=matches.at(-1)!;next=text.slice(0,first.index)+op.oneBasedOrder.map((i:number)=>matches[i-1]![0]).join('')+text.slice(last.index+last[0].length);}
  parts.set(op.part,encoder.encode(next));
 }
 return writeZip(parts);
}
