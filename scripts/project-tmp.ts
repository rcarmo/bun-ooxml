// Portable standalone resolver. Snapshot original TMPDIR before child redirection.
import {accessSync,constants,existsSync,lstatSync,mkdirSync,statSync} from 'node:fs';
import {basename,dirname,isAbsolute,join} from 'node:path';
import {tmpdir} from 'node:os';
export const project='bun-ooxml';
const originalTmpdir=process.env.TMPDIR,platformTmp=tmpdir();
function usable(path:string){
 if(!isAbsolute(path)||path.split(/[\\/]/).some(p=>p==='.'||p==='..'))return false;
 try{
  let ancestor=path;while(ancestor!==dirname(ancestor)){try{if(lstatSync(ancestor).isSymbolicLink()&&ancestor!=='/workspace')return false}catch(e){if((e as NodeJS.ErrnoException).code!=='ENOENT')return false};ancestor=dirname(ancestor)}
  if(existsSync(path)){const st=statSync(path);if(!st.isDirectory()||typeof process.getuid==='function'&&st.uid!==process.getuid())return false}
  let parent=path;while(!existsSync(parent))parent=dirname(parent);accessSync(parent,constants.W_OK|constants.X_OK);return true;
 }catch{return false}
}
export function resolveProjectTmp(env:NodeJS.ProcessEnv=process.env,workspace='/workspace/tmp',platform=platformTmp){
 const hasBase=Object.hasOwn(env,'PROJECT_TMP_BASE'),hasRoot=Object.hasOwn(env,'PROJECT_TMP_ROOT');
 if(hasBase||hasRoot){
  const base=env.PROJECT_TMP_BASE?.replace(/\/$/,'')??'',root=env.PROJECT_TMP_ROOT?.replace(/\/$/,'')??'';
  if(hasBase&&(!base||!usable(base)))throw Error('PROJECT_TMP_BASE must be a usable absolute directory');
  const candidate=hasBase?join(base,project):root;
  if(hasRoot&&(basename(root)!==project||!usable(root)||hasBase&&root!==candidate))throw Error('PROJECT_TMP_ROOT must be usable, project-named and agree with PROJECT_TMP_BASE');
  if(!usable(candidate))throw Error('Unsafe explicit project temporary directory');return candidate;
 }
 const ci=env.CI==='true'||env.CI==='1';
 const bases=ci?[env.RUNNER_TEMP,env===process.env?originalTmpdir:env.TMPDIR,platform]:[workspace,platform];
 for(const base of bases){if(!base)continue;const path=join(base,project);if(usable(path))return path}
 throw Error('No writable project-owned temp root');
}
export function initializeRoot(root:string){if(basename(root)!==project||!usable(root))throw Error('Unsafe project root');for(const child of ['cache','build','tests','logs','runs']){const path=join(root,child);if(!usable(path))throw Error('Unsafe project child');mkdirSync(path,{recursive:true})};return root}
if(import.meta.main){console.log(initializeRoot(resolveProjectTmp()))}
