// Development tooling only. Runtime packages must not import host path policy.
import {mkdirSync,realpathSync,lstatSync} from 'node:fs';
import {isAbsolute,join,resolve,relative,sep,dirname} from 'node:path';
import {resolveProjectTmp,initializeRoot} from './project-tmp.ts';
export function ownedRoot(){return initializeRoot(resolveProjectTmp());}
export function childPath(root:string,path:string){
 const full=resolve(root,path),rel=relative(root,full);if(!rel||rel==='..'||rel.startsWith('..'+sep)||isAbsolute(rel))throw Error('Unsafe project child path');
 let ancestor=full;while(ancestor!==root){try{if(lstatSync(ancestor).isSymbolicLink())throw Error('Project child cannot be symlink redirected')}catch(e){if((e as NodeJS.ErrnoException).code!=='ENOENT')throw e}ancestor=dirname(ancestor)}
 mkdirSync(full,{recursive:true});if(realpathSync(full)!==join(realpathSync(root),rel))throw Error('Project child cannot be symlink redirected');return full;
}
export function prepareDevPaths(){
 const root=ownedRoot();const run=process.env.OOXML_RUN_DIR??childPath(root,join('runs','direct',`${process.pid}-${crypto.randomUUID()}`));
 const rel=relative(root,run);if(!rel.startsWith('runs'+sep)||rel.includes('..')||!isAbsolute(run))throw Error('Run must be under project runs');
 const scratch=realpathSync(childPath(root,join(rel,'tmp')));
 process.env.PROJECT_TMP_ROOT=root;process.env.OOXML_TMP_ROOT=root;process.env.OOXML_RUN_DIR=run;
 process.env.PROJECT_TEST_ROOT=childPath(root,'tests');process.env.PROJECT_LOG_ROOT=childPath(root,'logs');
 for(const key of ['TMPDIR','TMP','TEMP'])process.env[key]=scratch;
 const paths:Record<string,string>={BUN_INSTALL_CACHE_DIR:'cache/bun',XDG_CACHE_HOME:'cache/xdg',NPM_CONFIG_CACHE:'cache/npm',DOTNET_CLI_HOME:'cache/dotnet-home',NUGET_PACKAGES:'cache/nuget',PYTHONPYCACHEPREFIX:'cache/python',OOXML_BUILD_DIR:'build'};
 for(const [key,path]of Object.entries(paths))process.env[key]=childPath(root,path);
 process.env.DOTNET_CLI_TELEMETRY_OPTOUT='1';process.env.DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1';
 return {root,run,scratch,build:process.env.OOXML_BUILD_DIR!};
}
