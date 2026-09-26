import {parse} from '@babel/parser';
import {join,relative,resolve} from 'node:path';
export type TestCaseRecord={id:string;path:string;line:number;title:string;suites:string[];kind:string;assertions:string[];loops:{line:number;expression:string}[];parameters?:string;sourceSha256:string;body:string};
const digest=(value:string)=>new Bun.CryptoHasher('sha256').update(value).digest('hex');
function callName(node:any):string{
 if(node?.type==='Identifier')return node.name;
 if(node?.type==='MemberExpression')return callName(node.object)+'.'+(node.computed?node.property.value:node.property.name);
 if(node?.type==='CallExpression')return callName(node.callee)+'()';
 return '';
}
export function inventoryTestSource(path:string,source:string):TestCaseRecord[]{
 const ast=parse(source,{sourceType:'module',plugins:['typescript']});const records:TestCaseRecord[]=[];
 const slice=(n:any)=>source.slice(n.start,n.end);
 function walk(n:any,suites:string[]){
  if(!n||typeof n!=='object')return;
  if(n.type==='CallExpression'){
   const name=callName(n.callee),callback=n.arguments.find((a:any)=>['ArrowFunctionExpression','FunctionExpression'].includes(a.type));
   const title=n.arguments[0]?.type==='StringLiteral'?n.arguments[0].value:n.arguments[0]?.type==='TemplateLiteral'?slice(n.arguments[0]):undefined;
   if(/^describe(?:\.|$)/.test(name)&&callback){walk(callback,[...suites,title??'<dynamic suite>']);return;}
   if(/^(?:test|it)(?:\.|$)/.test(name)&&callback){
    const assertions:string[]=[],loops:{line:number;expression:string}[]=[];
    function inspect(x:any){if(!x||typeof x!=='object')return;
     if(x.type==='CallExpression'&&/^(?:expect\(\)|assert(?:\.|\())/.test(callName(x.callee)))assertions.push(slice(x));
     if(['ForOfStatement','ForInStatement','ForStatement','WhileStatement'].includes(x.type))loops.push({line:x.loc.start.line,expression:slice(x).split('{')[0]!.trim()});
     for(const[k,v]of Object.entries(x))if(!['loc','start','end','extra','comments','tokens'].includes(k)){if(Array.isArray(v))v.forEach(inspect);else if(v&&typeof v==='object')inspect(v);}
    }inspect(callback.body);
    const id='bun:'+path+':'+[...suites,title??'<dynamic>'].join(' / ');
    records.push({id,path,line:n.loc.start.line,title:title??'<dynamic>',suites,kind:name,assertions:[...new Set(assertions)],loops,sourceSha256:digest(source),body:slice(callback.body),...(n.callee.type==='CallExpression'?{parameters:slice(n.callee.arguments[0])}:{})});return;
   }
  }
  for(const[k,v]of Object.entries(n))if(!['loc','start','end','extra','comments','tokens'].includes(k)){if(Array.isArray(v))for(const c of v)walk(c,suites);else if(v&&typeof v==='object')walk(v,suites);}
 }
 walk(ast,[]);return records;
}
export async function inventoryNativeTests(root=resolve(import.meta.dir,'..')){
 const cases:TestCaseRecord[]=[];for await(const path of new Bun.Glob('tests/unit/**/*.test.ts').scan({cwd:root}))cases.push(...inventoryTestSource(path,await Bun.file(join(root,path)).text()));
 cases.sort((a,b)=>a.path.localeCompare(b.path)||a.line-b.line);
 if(new Set(cases.map(c=>c.id)).size!==cases.length)throw Error('Duplicate native test identities');
 return {schemaVersion:1,consumer:'bun',discovery:'TypeScript AST; declarations are not runtime-expanded leaf cases',count:cases.length,files:new Set(cases.map(c=>c.path)).size,dynamicCases:cases.filter(c=>c.title.includes('${')||c.parameters||c.loops.length).map(c=>c.id),cases};
}
if(import.meta.main){const inventory=await inventoryNativeTests();await Bun.write(join(import.meta.dir,'../docs/behaviors/native-test-inventory.json'),JSON.stringify(inventory,null,2)+'\n');console.log(`${inventory.count} declarations in ${inventory.files} files; ${inventory.dynamicCases.length} parameter/loop cases need explicit review`);}
