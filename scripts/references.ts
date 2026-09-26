import {join} from 'node:path';
const digest=(bytes:Uint8Array)=>new Bun.CryptoHasher('sha256').update(bytes).digest('hex');
const safe=(path:unknown):path is string=>typeof path==='string'&&!!path&&!/[\\:\u0000-\u001f]/.test(path)&&path.split('/').every(p=>!!p&&p!=='.'&&p!=='..');
export async function verifyReferences(root:string):Promise<number>{
 const refs=join(root,'references/fixtures-ooxml'),pinFile=Bun.file(join(root,'references/fixtures-ooxml.pin.json'));
 if(!await pinFile.exists())throw Error('Missing required shared reference pin');
 const pin=await pinFile.json();
 if(pin.repository!=='https://github.com/rcarmo/fixtures-ooxml'||!/^v\d+\.\d+\.\d+$/.test(pin.tag)||! /^[a-f0-9]{40}$/.test(pin.commit))throw Error('Invalid shared reference pin');
 const git=(...args:string[])=>{const r=Bun.spawnSync(['git','-C',refs,...args]);if(r.exitCode)throw Error('Missing shared submodule: git submodule update --init --recursive');return r.stdout.toString().trim();};
 if(git('rev-parse','HEAD')!==pin.commit||git('rev-parse',`refs/tags/${pin.tag}^{commit}`)!==pin.commit||git('cat-file','-t',`refs/tags/${pin.tag}`)!=='tag')throw Error('Shared submodule/tag pin mismatch');
 const manifestBytes=await Bun.file(join(refs,'manifest.json')).bytes();
 if(digest(manifestBytes)!==pin.manifestSha256)throw Error('Shared manifest seal mismatch');
 if(digest(await Bun.file(join(refs,'shared/v2/pack/pack-manifest.json')).bytes())!==pin.sharedPackManifestSha256)throw Error('Shared pack seal mismatch');
 const manifest=JSON.parse(new TextDecoder().decode(manifestBytes)),known=new Set<string>();
 if(manifest.schemaVersion!==1||!Array.isArray(manifest.files)||!manifest.files.length)throw Error('Invalid shared manifest');
 for(const f of manifest.files){if(!safe(f.path)||known.has(f.path)||!Number.isSafeInteger(f.bytes)||f.bytes<0||!/^[a-f0-9]{64}$/.test(f.sha256))throw Error('Invalid reference path or hash');known.add(f.path);const bytes=await Bun.file(join(refs,f.path)).bytes();if(bytes.length!==f.bytes||digest(bytes)!==f.sha256)throw Error('Reference drift: '+f.path);}
 return known.size;
}
