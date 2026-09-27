import {join,resolve} from 'node:path';
import {inventoryFeatures} from './gherkin.ts';
import {verifyReferences} from './references.ts';
const root=resolve(import.meta.dir,'..');
const count=await verifyReferences(root);
for await(const path of new Bun.Glob('src/**/*.ts').scan({cwd:root})){const source=await Bun.file(join(root,path)).text();if(/node:child_process|bun:ffi|Bun\.(?:spawn|dlopen)|\b(?:execFile|spawnSync)\s*\(/.test(source))throw Error('Foreign runtime bridge: '+path);}
if(Object.keys((await Bun.file(join(root,'package.json')).json()).dependencies??{}).length)throw Error('Runtime dependencies need review');
const acceptance=await inventoryFeatures(root);
if(!acceptance.coverage?.catalogue)throw Error('Complete shared catalogue inventory is required for the project');
if(process.argv.includes('--full')&&acceptance.counts.cases.planned>0)throw Error('Full behaviour coverage incomplete: planned contracts remain');
console.log(`Verified ${count} shared assets; ${acceptance.coverage.shared.scenarios.total} shared scenarios / ${acceptance.coverage.shared.cases.total} cases, plus ${acceptance.coverage.localOnly.scenarios.total} local-only scenarios / ${acceptance.coverage.localOnly.cases.total} cases; no runtime dependencies.`);
