import {join,resolve} from 'node:path';
import {inventoryFeatures} from './gherkin.ts';
const root=resolve(import.meta.dir,'..'),refs=join(root,'references/fixtures-ooxml');
const file=Bun.file(join(refs,'manifest.json'));
if(!await file.exists())throw Error('Missing shared references: git submodule update --init --recursive');
const manifest=await file.json();let count=0;const known=new Set<string>();
for(const f of manifest.files){if(typeof f.path!=='string'||f.path.split('/').some((x:string)=>!x||x==='..'||x==='.')||known.has(f.path))throw Error('Invalid reference path');known.add(f.path);const bytes=await Bun.file(join(refs,f.path)).bytes();if(bytes.length!==f.bytes||new Bun.CryptoHasher('sha256').update(bytes).digest('hex')!==f.sha256)throw Error('Reference drift: '+f.path);count++;}
const pinFile=Bun.file(join(root,'references/fixtures-ooxml.pin.json'));
if(await pinFile.exists()){
 const pin=await pinFile.json(),result=Bun.spawnSync(['git','-C',refs,'rev-parse','HEAD']);
 if(result.exitCode||result.stdout.toString().trim()!==pin.commit)throw Error('Shared submodule pin mismatch');
}
for await(const path of new Bun.Glob('src/**/*.ts').scan({cwd:root})){const source=await Bun.file(join(root,path)).text();if(/node:child_process|bun:ffi|Bun\.(?:spawn|dlopen)|\b(?:execFile|spawnSync)\s*\(/.test(source))throw Error('Foreign runtime bridge: '+path);}
if(Object.keys((await Bun.file(join(root,'package.json')).json()).dependencies??{}).length)throw Error('Runtime dependencies need review');
const acceptance=await inventoryFeatures(root);
if(process.argv.includes('--full')&&acceptance.counts.cases.planned>0)throw Error('Full behaviour coverage incomplete: planned contracts remain');
console.log(`Verified ${count} shared assets; ${acceptance.counts.scenarios.total} local/shared scenarios; no runtime dependencies.`);
