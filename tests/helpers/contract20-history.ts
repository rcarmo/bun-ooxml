import assert from 'node:assert/strict';
import pin from '../../docs/behaviors/contract20-candidate.json';
import {fixturesRoot} from '../../scripts/fixture-inputs.ts';
export async function beforeContract20Feature(path:string,text:string):Promise<string>{const ledger=Bun.file(fixturesRoot()+'/ledgers/contract20.json');if(!await ledger.exists())return text;const data=await ledger.json(),f=data.files.find((f:any)=>f.path===path);if(!f||text===f.beforeText)return text;assert.equal(new Bun.CryptoHasher('sha256').update(text).digest('hex'),f.afterSha256,'Unreviewed Contract20 historical feature');assert.equal(pin.featureSeals[path as keyof typeof pin.featureSeals],f.afterSha256);return f.beforeText;}
