/** @script Discover native Bun test declarations and check the committed inventory.
 * @usage bun scripts/test-inventory.ts [--check]
 * @description AST discovery only; does not expand runtime cases or award execution credit.
 */
import {parse} from '@babel/parser';
import {join,resolve} from 'node:path';
export type TestCaseRecord={id:string;path:string;line:number;title:string;suites:string[];kind:string;lifecycle:string;assertions:string[];deferredAssertions:string[];loops:{line:number;expression:string}[];registrationLoops:{line:number;expression:string}[];parameters?:string;sourceSha256:string;body:string;calls:string[];reviewReasons:string[];unresolved:string[]};
const digest=(value:string)=>new Bun.CryptoHasher('sha256').update(value).digest('hex');
const skip=new Set(['loc','start','end','extra','comments','tokens']);
const loopTypes=new Set(['ForOfStatement','ForInStatement','ForStatement','WhileStatement','DoWhileStatement']);
const callback=(n:any)=>n&&['ArrowFunctionExpression','FunctionExpression'].includes(n.type);
function children(n:any,fn:(n:any)=>void){for(const[k,v]of Object.entries(n))if(!skip.has(k)){if(Array.isArray(v))v.forEach(fn);else if(v&&typeof v==='object')fn(v);}}
function unwrap(n:any):any {return n&&['TSAsExpression','TSTypeAssertion','TSNonNullExpression','TSSatisfiesExpression','ParenthesizedExpression'].includes(n.type)?unwrap(n.expression):n;}
function callName(n:any):string{
 n=unwrap(n);
 if(n?.type==='Identifier')return n.name;
 if(n?.type==='MemberExpression')return callName(n.object)+'.'+(n.computed?n.property.value:n.property.name);
 if(n?.type==='CallExpression')return callName(n.callee)+'()';
 return '';
}
type Context={suites:string[];reasons:string[];loops:{line:number;expression:string}[];deferred?:boolean};
function bindings(n:any):string[]{
 if(!n)return [];if(n.type==='Identifier')return[n.name];if(n.type==='AssignmentPattern')return bindings(n.left);if(n.type==='RestElement')return bindings(n.argument);
 if(n.type==='ObjectPattern')return n.properties.flatMap((p:any)=>bindings(p.type==='RestElement'?p.argument:p.value));
 if(n.type==='ArrayPattern')return n.elements.flatMap(bindings);return [];
}
export function inventoryTestSource(path:string,source:string):TestCaseRecord[]{return analyse(path,source).records;}
function analyse(path:string,source:string){
 const ast=parse(source,{sourceType:'module',plugins:['typescript']}),records:TestCaseRecord[]=[],issues:string[]=[];
 const aliases=new Map<string,string>(),namespaces=new Set<string>(),slice=(n:any)=>n?source.slice(n.start,n.end):'';
 for(const statement of ast.program.body)if(statement.type==='ImportDeclaration'&&statement.source.value==='bun:test')for(const s of statement.specifiers){
  if(s.type==='ImportNamespaceSpecifier')namespaces.add(s.local.name);
  else if(s.type==='ImportSpecifier')aliases.set(s.local.name,s.imported.type==='Identifier'?s.imported.name:s.imported.value);
  else issues.push(`line ${s.loc?.start.line}: unsupported default bun:test import`);
 }
 function canonical(n:any):string{
  n=unwrap(n);
  if(n?.type==='Identifier')return aliases.get(n.name)??'';
  if(n?.type==='MemberExpression'){
   const property=n.computed?n.property.type==='StringLiteral'?n.property.value:'<dynamic>':n.property.name;
   if(n.object.type==='Identifier'&&namespaces.has(n.object.name))return property;
   const base=canonical(n.object);return base?base+'.'+property:'';
  }
  if(n?.type==='CallExpression'){const base=canonical(n.callee);return base?base+'()':'';}
  return '';
 }
 function title(n:any){return n?.type==='StringLiteral'?n.value:n?.type==='TemplateLiteral'&&!n.expressions.length?n.quasis[0].value.cooked:undefined;}
 function loop(n:any){return {line:n.loc.start.line,expression:slice(n).split('{')[0]!.trim()};}
 // Registration aliases and binding shadowing need scope evaluation. Flag them instead
 // of silently counting an imported test name that no longer refers to Bun.
 function audit(n:any){
  if(!n||typeof n!=='object')return;
  if(n.type==='VariableDeclarator'){
   if(canonical(n.init)||unwrap(n.init)?.type==='Identifier'&&namespaces.has(unwrap(n.init).name))issues.push(`line ${n.loc.start.line}: indirect bun:test alias requires review`);
   if(bindings(n.id).some(id=>aliases.has(id)||namespaces.has(id)))issues.push(`line ${n.loc.start.line}: shadowed bun:test binding`);
  }
  if(['FunctionDeclaration','FunctionExpression','ArrowFunctionExpression'].includes(n.type))for(const p of n.params??[])if(bindings(p).some(id=>aliases.has(id)||namespaces.has(id)))issues.push(`line ${p.loc.start.line}: shadowed bun:test parameter`);
  if(n.type==='CatchClause'&&bindings(n.param).some(id=>aliases.has(id)||namespaces.has(id)))issues.push(`line ${n.loc.start.line}: shadowed bun:test catch binding`);
  children(n,audit);
 }
 audit(ast);
 function walk(n:any,ctx:Context){
  if(!n||typeof n!=='object')return;
  if(['FunctionDeclaration','FunctionExpression','ArrowFunctionExpression'].includes(n.type)){children(n,c=>walk(c,{...ctx,deferred:true}));return;}
  if(loopTypes.has(n.type)){children(n,c=>walk(c,{...ctx,reasons:[...ctx.reasons,'registration-loop'],loops:[...ctx.loops,loop(n)]}));return;}
  if(['IfStatement','ConditionalExpression','LogicalExpression'].includes(n.type)){children(n,c=>walk(c,{...ctx,reasons:[...ctx.reasons,'conditional-registration']}));return;}
  if(n.type==='CallExpression'){
   const name=canonical(n.callee),fn=n.arguments.find(callback),literal=title(n.arguments[0]),label=literal??(slice(n.arguments[0])||'<dynamic>');
   if(/^describe(?:\.|$)/.test(name)){
    if(ctx.deferred)issues.push(`line ${n.loc.start.line}: registration inside non-suite function requires review`);
    if(!fn){issues.push(`line ${n.loc.start.line}: non-inline describe callback`);return;}
    const reasons=[...ctx.reasons,...(literal===undefined?['dynamic-suite']:[]),...(name.includes('.each')?['parameterized-suite']:[]),...(name!=='describe'?['suite-modifier']:[])];
    walk(fn.body,{...ctx,suites:[...ctx.suites,label],reasons});return;
   }
   if(/^(?:test|it)(?:\.|$)/.test(name)){
    const assertions:string[]=[],deferredAssertions:string[]=[],loops:{line:number;expression:string}[]=[],calls=new Set<string>(),unresolved:string[]=[];
    const todo=name.split('.').includes('todo'),lifecycle=todo?'todo':name.split('.').includes('skip')?'skip':name.split('.').includes('only')?'only':'normal';
    if(ctx.deferred)unresolved.push('registration-inside-non-suite-function');
    if(!fn&&!todo)unresolved.push('non-inline-callback');
    if(name.includes('<dynamic>'))unresolved.push('dynamic-registration-modifier');
    function inspect(x:any,deferred=false){
     if(!x||typeof x!=='object')return;
     const nested=deferred||['FunctionDeclaration','FunctionExpression','ArrowFunctionExpression'].includes(x.type);
     if(x.type==='CallExpression'){
      const called=canonical(x.callee),raw=callName(x.callee);calls.add(raw);
      if(/^expect\(\)\./.test(called))(nested?deferredAssertions:assertions).push(slice(x));
     }
     if(loopTypes.has(x.type))loops.push(loop(x));
     children(x,c=>inspect(c,nested));
    }
    if(fn)inspect(fn.body);
    const parameters=n.callee.type==='CallExpression'&&name.includes('.each')?slice(n.callee.arguments[0]):undefined;
    const reasons=[...ctx.reasons,...(/\.(?:if|skipIf)\(/.test(name)?['conditional-registration']:[]),...(name.includes('.')?['test-modifier']:[]),...(literal===undefined?['dynamic-title']:[]),...(parameters!==undefined?['parameterized-test']:[]),...(loops.length?['body-loop']:[]),...(deferredAssertions.length?['deferred-assertion']:[]),...(!assertions.length?['helper-or-no-direct-assertion']:[]),...(lifecycle!=='normal'?[`lifecycle-${lifecycle}`]:[])];
    records.push({id:'bun:'+path+':'+[...ctx.suites,label].join(' / '),path,line:n.loc.start.line,title:label,suites:ctx.suites,kind:name,lifecycle,assertions:[...new Set(assertions)],deferredAssertions:[...new Set(deferredAssertions)],loops,registrationLoops:ctx.loops,sourceSha256:digest(source),body:fn?slice(fn.body):'',calls:[...calls].filter(Boolean).sort(),reviewReasons:[...new Set(reasons)].sort(),unresolved,...(parameters!==undefined?{parameters}:{})});return;
   }
  }
  if(n.type==='TaggedTemplateExpression'&&canonical(n.tag)){issues.push(`line ${n.loc.start.line}: tagged bun:test declaration/table requires review`);return;}
  children(n,c=>walk(c,ctx));
 }
 walk(ast,{suites:[],reasons:[],loops:[]});return {records,issues:[...new Set(issues)]};
}
export async function inventoryNativeTests(root=resolve(import.meta.dir,'..')){
 const testSupport:{path:string;sha256:string}[]=[];
 for await(const path of new Bun.Glob('tests/**/*.ts').scan({cwd:root}))if(!/^tests\/unit\/.*\.test\.ts$/.test(path))testSupport.push({path,sha256:digest(await Bun.file(join(root,path)).text())});
 testSupport.sort((a,b)=>a.path.localeCompare(b.path));
 const cases:TestCaseRecord[]=[],sources:{path:string;sha256:string;declarations:number;issues:string[]}[]=[];
 for await(const path of new Bun.Glob('tests/unit/**/*.test.ts').scan({cwd:root})){
  const source=await Bun.file(join(root,path)).text(),result=analyse(path,source);cases.push(...result.records);sources.push({path,sha256:digest(source),declarations:result.records.length,issues:result.issues});
 }
 sources.sort((a,b)=>a.path.localeCompare(b.path));cases.sort((a,b)=>a.path.localeCompare(b.path)||a.line-b.line);
 if(new Set(cases.map(c=>c.id)).size!==cases.length)throw Error('Duplicate native test identities');
 return {schemaVersion:2,consumer:'bun',discovery:'TypeScript AST of named/namespace bun:test imports; declarations are not runtime-expanded leaf cases',executionCredit:false,count:cases.length,files:sources.length,sources,testSupport,dynamicCases:cases.filter(c=>c.reviewReasons.some(r=>!['helper-or-no-direct-assertion','deferred-assertion'].includes(r))).map(c=>c.id),reviewRequired:cases.filter(c=>c.reviewReasons.length||c.unresolved.length).map(c=>c.id),unresolved:[...sources.flatMap(s=>s.issues.map(issue=>({path:s.path,issue}))),...cases.flatMap(c=>c.unresolved.map(issue=>({path:c.path,issue:`line ${c.line}: ${issue}`})))],cases};
}
export async function checkInventory(root=resolve(import.meta.dir,'..'),reportPath=join(root,'docs/behaviors/native-test-inventory.json')):Promise<void>{
 const inventory=await inventoryNativeTests(root);if(inventory.unresolved.length)throw Error('Unresolved native registrations: '+JSON.stringify(inventory.unresolved));
 if(!await Bun.file(reportPath).exists()||await Bun.file(reportPath).text()!==JSON.stringify(inventory,null,2)+'\n')throw Error('Native test inventory stale; run bun scripts/test-inventory.ts');
}
if(import.meta.main){const root=resolve(import.meta.dir,'..'),inventory=await inventoryNativeTests(root);if(process.argv.includes('--check'))await checkInventory(root);else{if(inventory.unresolved.length)throw Error('Unresolved native registrations: '+JSON.stringify(inventory.unresolved));await Bun.write(join(root,'docs/behaviors/native-test-inventory.json'),JSON.stringify(inventory,null,2)+'\n');}console.log(`${inventory.count} declarations in ${inventory.files} files; ${inventory.dynamicCases.length} dynamic declarations; ${inventory.reviewRequired.length} need assertion/expansion review; no execution credit`);}
