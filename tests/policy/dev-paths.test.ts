import {test,expect} from 'bun:test';
import {mkdtempSync,rmSync,symlinkSync,mkdirSync,realpathSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join,relative} from 'node:path';
import {childPath,prepareDevPaths} from '../../scripts/dev-paths.ts';
import {resolveProjectTmp} from '../../scripts/project-tmp.ts';
test('development cache and scratch paths remain project owned',()=>{
 const paths=prepareDevPaths();expect(paths.root.endsWith('/bun-ooxml')).toBe(true);
 expect(tmpdir()).toBe(paths.scratch);
 for(const key of ['TMPDIR','TMP','TEMP'])expect(process.env[key]).toBe(paths.scratch);
 for(const key of ['BUN_INSTALL_CACHE_DIR','XDG_CACHE_HOME','NPM_CONFIG_CACHE','NUGET_PACKAGES','DOTNET_CLI_HOME','PYTHONPYCACHEPREFIX'])expect(process.env[key]?.startsWith(paths.root+'/cache/')).toBe(true);
 expect(paths.build).toBe(paths.root+'/build');
});
test('portable explicit BASE/ROOT and CI/local precedence retain project hierarchy',()=>{
 const paths=prepareDevPaths(),dir=mkdtempSync(join(paths.scratch,'fallback-'));
 try{const blocked=join(dir,'file');require('node:fs').writeFileSync(blocked,'blocked');
 expect(()=>resolveProjectTmp({PROJECT_TMP_ROOT:'relative/bun-ooxml'},blocked,dir)).toThrow();
 expect(()=>resolveProjectTmp({PROJECT_TMP_BASE:''},blocked,dir)).toThrow();
 const base=join(dir,'override'),explicit=join(base,'bun-ooxml');
 expect(resolveProjectTmp({PROJECT_TMP_ROOT:explicit},blocked,dir)).toBe(explicit);
 expect(resolveProjectTmp({PROJECT_TMP_BASE:base},blocked,dir)).toBe(explicit);
 expect(resolveProjectTmp({PROJECT_TMP_BASE:base,PROJECT_TMP_ROOT:explicit},blocked,dir)).toBe(explicit);
 expect(()=>resolveProjectTmp({PROJECT_TMP_BASE:base,PROJECT_TMP_ROOT:join(dir,'other','bun-ooxml')},blocked,dir)).toThrow();
 expect(resolveProjectTmp({CI:'true',RUNNER_TEMP:dir,TMPDIR:join(dir,'inherited')},join(dir,'workspace'),join(dir,'platform'))).toBe(join(dir,'bun-ooxml'));
 expect(resolveProjectTmp({CI:'true',TMPDIR:dir},join(dir,'workspace'),join(dir,'platform'))).toBe(join(dir,'bun-ooxml'));
 expect(resolveProjectTmp({RUNNER_TEMP:join(dir,'runner'),TMPDIR:join(dir,'inherited')},blocked,dir)).toBe(join(dir,'bun-ooxml'));
 expect(resolveProjectTmp({},dir,join(dir,'platform'))).toBe(join(dir,'bun-ooxml'));
 }finally{rmSync(dir,{recursive:true,force:true})}
});
test('development child paths refuse traversal and symlink redirection',()=>{
 const paths=prepareDevPaths(),owned=mkdtempSync(join(paths.scratch,'paths-'));
 try{expect(()=>childPath(paths.root,'../other')).toThrow();const outside=join(owned,'outside');mkdirSync(outside);const link=join(owned,'link');symlinkSync(outside,link);expect(()=>childPath(paths.root,relative(realpathSync(paths.root),link))).toThrow();}
 finally{rmSync(owned,{recursive:true,force:true})}
});
