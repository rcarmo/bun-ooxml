import assert from 'node:assert/strict';
import {beforeContract20Feature} from './contract20-history.ts';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
import {join} from 'node:path';
/** Test-only exact strengthening reversal for historical bindings/fingerprints. */
export async function beforeUniformApi18Feature(path:string,text:string):Promise<string>{
 text=await beforeContract20Feature(path,text);
 const file=Bun.file(join(fixturesRoot(),'ledgers/uniform-api18.json'));if(!await file.exists())return text;
 const ledger=await file.json(),f=ledger.files.find((f:any)=>f.path===path);if(!f||text===f.beforeText)return text;
 assert.equal(new Bun.CryptoHasher('sha256').update(text).digest('hex'),f.afterSha256,'Unreviewed API18 feature reversal');return f.beforeText;
}
