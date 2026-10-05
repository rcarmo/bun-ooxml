// Related bounded test processes avoid retaining all catalogue ASTs in one heap.
// Keep every test file/assertion/timeout; profile batches only when opted in.
import {resolve} from 'node:path';
const root=resolve(import.meta.dir,'..');
const files=[...new Bun.Glob('tests/unit/**/*.test.ts').scanSync({cwd:root})].sort();
const harness=files.filter(p=>/inventory|mappings?|outcomes?|gherkin|references|coverage/.test(p));
const runtime=files.filter(p=>!harness.includes(p)&&p!=='tests/unit/oracle-process.test.ts');
const batches:Array<{label:string;files:string[]}>=[];
for(let i=0;i<harness.length;i+=2)batches.push({label:`catalogue-${i/2+1}`,files:harness.slice(i,i+2)});
batches.push({label:'unit-runtime',files:runtime}, {label:'unit-policy',files:[...new Bun.Glob('tests/policy/**/*.test.ts').scanSync({cwd:root}), 'tests/unit/oracle-process.test.ts']});
const seen=new Set(batches.flatMap(b=>b.files));if(files.some(f=>!seen.has(f)))throw Error('Incomplete unit batch inventory');
const from=process.argv[2];if(from&&!batches.some(b=>b.label===from))throw Error('Unknown resume batch');
if(from)console.log(`Resuming from ${from}; earlier concise batch conclusions require separate verification`);
for(const batch of batches.slice(from?batches.findIndex(b=>b.label===from):0)){const interval=batch.label==='unit-policy'?'1000':batch.files.some(p=>/comment-thread|revision-move/.test(p))?'100000':'10000';const child=Bun.spawn([process.execPath,resolve(root,'scripts/dev-run.ts'),batch.label,'test',...batch.files],{cwd:root,env:{...process.env,OOXML_CPU_INTERVAL_US:interval},stdout:'inherit',stderr:'inherit'});const code=await child.exited;if(code){process.exitCode=code;break}}
