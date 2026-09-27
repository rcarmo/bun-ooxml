import {join,resolve} from 'node:path';
import {lstat,readFile,realpath} from 'node:fs/promises';
const PROJECT_ROOT=resolve(import.meta.dir,'..');
const digest=(bytes:Uint8Array)=>new Bun.CryptoHasher('sha256').update(bytes).digest('hex');
const safe=(path:unknown):path is string=>typeof path==='string'&&!!path&&!/[\\:\u0000-\u001f]/.test(path)&&path.split('/').every(p=>!!p&&p!=='.'&&p!=='..');
export async function verifyReferences(root:string):Promise<number>{
 const isProject=resolve(root)===PROJECT_ROOT;
 const refs=isProject&&process.env.OOXML_FIXTURES_ROOT?resolve(process.env.OOXML_FIXTURES_ROOT):join(root,'references/fixtures-ooxml');
 const pinFile=Bun.file(isProject&&process.env.OOXML_REFERENCE_PIN?process.env.OOXML_REFERENCE_PIN:join(root,'references/fixtures-ooxml.pin.json'));
 if(isProject&&process.env.OOXML_FIXTURES_ROOT&&!process.env.OOXML_REFERENCE_PIN)throw Error('Candidate fixture root requires explicit candidate pin');
 if(!await pinFile.exists())throw Error('Missing required shared reference pin');
 const pin=await pinFile.json();
 const candidate=pin.mode==='candidate';
 if(candidate&&pin.tag!==undefined)throw Error('Candidate pin must not claim a release tag');
 if(candidate&&isProject&&(!process.env.OOXML_FIXTURES_ROOT||!process.env.OOXML_REFERENCE_PIN))throw Error('Candidate mode requires root and pin overrides');
 if(pin.repository!=='https://github.com/rcarmo/fixtures-ooxml'||(!candidate&&!/^v\d+\.\d+\.\d+$/.test(pin.tag))||! /^[a-f0-9]{40}$/.test(pin.commit))throw Error('Invalid shared reference pin');
 const git=(...args:string[])=>{const r=Bun.spawnSync(['git','-C',refs,...args]);if(r.exitCode)throw Error('Missing shared submodule: git submodule update --init --recursive');return r.stdout.toString().trim();};
 if(git('rev-parse','HEAD')!==pin.commit||(!candidate&&(git('rev-parse',`refs/tags/${pin.tag}^{commit}`)!==pin.commit||git('cat-file','-t',`refs/tags/${pin.tag}`)!=='tag')))throw Error('Shared submodule/tag pin mismatch');
 const manifestBytes=await Bun.file(join(refs,'manifest.json')).bytes();
 if(digest(manifestBytes)!==pin.manifestSha256)throw Error('Shared manifest seal mismatch');
 const manifest=JSON.parse(new TextDecoder().decode(manifestBytes)),known=new Set<string>();
 if(manifest.schemaVersion!==2||manifest.fixturePathBase!=='repository-root'||!Array.isArray(manifest.files)||!manifest.files.length)throw Error('Invalid shared manifest');
 const hashes=new Set<string>(),ids=new Set<string>();
 for(const f of manifest.files){if(!safe(f.path)||known.has(f.path)||hashes.has(f.sha256)||ids.has(f.id)||!Number.isSafeInteger(f.bytes)||f.bytes<0||!/^[a-f0-9]{64}$/.test(f.sha256))throw Error('Invalid or duplicate reference path/hash/ID');known.add(f.path);hashes.add(f.sha256);ids.add(f.id);
 if(f.role==='fixture'&&(f.id!=='fixture-'+f.sha256||!['docx','pptx','xlsx','png','zip'].includes(f.format)||!/^[a-z0-9-]+$/.test(f.scenarioGroup)||!f.path.startsWith(`fixtures/${f.format}/${f.scenarioGroup}/`)))throw Error('Invalid grouped fixture identity');const bytes=await Bun.file(join(refs,f.path)).bytes();if(bytes.length!==f.bytes||digest(bytes)!==f.sha256)throw Error('Reference drift: '+f.path);}
 if(git('--no-optional-locks','status','--porcelain','--untracked-files=all'))throw Error('Shared reference worktree is dirty');
 // Status honours index hints/filters. Compare raw tracked bytes independently.
 const objectFormat=git('rev-parse','--show-object-format');
 if(objectFormat!=='sha1'&&objectFormat!=='sha256')throw Error('Unsupported reference object format');
 const referenceRoot=await realpath(refs);
 for(const record of git('ls-tree','-r','-z','--full-tree','HEAD').split('\0').filter(Boolean)){
  const match=record.match(/^(100644|100755) blob ([a-f0-9]+)\t([^\0]+)$/);
  if(!match||!safe(match[3]))throw Error('Unsupported tracked reference entry');
  const [,mode,objectId,path]=match,full=join(referenceRoot,path!);
  const stat=await lstat(full);
  if(!stat.isFile()||await realpath(full)!==full)throw Error('Tracked reference must be a regular non-symlink file: '+path);
  if(process.platform!=='win32'&&Boolean(stat.mode&0o111)!==(mode==='100755'))throw Error('Tracked reference mode differs: '+path);
  const bytes=await readFile(full),actual=new Bun.CryptoHasher(objectFormat).update(`blob ${bytes.length}\0`).update(bytes).digest('hex');
  if(actual!==objectId)throw Error('Tracked reference bytes differ: '+path);
 }
 return known.size;
}
