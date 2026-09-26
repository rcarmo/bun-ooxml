import {test,expect} from 'bun:test';
import {inventoryTestSource,inventoryNativeTests,checkInventory} from '../../scripts/test-inventory.ts';
import {mkdtemp,mkdir,rm} from 'node:fs/promises';
import {join} from 'node:path';
import {tmpdir} from 'node:os';

test('native inventory recognises aliased and namespace Bun registrations without unrelated lookalikes',()=>{
 const source=`import {test as spec, describe as suite, expect as check} from 'bun:test';import * as bt from 'bun:test';suite('suite',()=>{spec('one',()=>{check(1).toBe(1)});bt.it('two',()=>{bt.expect(2).toBe(2)})});other.test('fake',()=>{});`;
 const rows=inventoryTestSource('x.test.ts',source);expect(rows).toHaveLength(2);expect(rows.map(r=>r.title)).toEqual(['one','two']);expect(rows[0]!.suites).toEqual(['suite']);expect(rows[0]!.assertions).toEqual(['check(1).toBe(1)']);expect(rows[1]!.assertions).toEqual(['bt.expect(2).toBe(2)']);
});
test('native inventory keeps declaration counts distinct from each tables and registration loops',()=>{
 const source=`import {test,describe} from 'bun:test';for(const v of [1,2]){test(\`row \${v}\`,()=>{})}describe.each([[1],[2]])('suite %i',()=>{test.each([[3],[4]])('item %i',n=>{})});`;
 const rows=inventoryTestSource('x.test.ts',source);expect(rows).toHaveLength(2);expect(rows[0]!.reviewReasons).toContain('dynamic-title');expect(rows[0]!.reviewReasons).toContain('registration-loop');expect(rows[1]!.parameters).toBe('[[3],[4]]');expect(rows[1]!.reviewReasons).toContain('parameterized-suite');expect(rows[1]!.reviewReasons).toContain('parameterized-test');
});
test('native inventory marks dynamic suite names, callbacks and conditional registrations explicitly',()=>{
 const source=`import {test,describe} from 'bun:test';describe(name,()=>{if(flag)test('conditional',()=>{})});test('opaque',handler);`;
 const rows=inventoryTestSource('x.test.ts',source);expect(rows).toHaveLength(2);expect(rows[0]!.reviewReasons).toContain('dynamic-suite');expect(rows[0]!.reviewReasons).toContain('conditional-registration');expect(rows[1]!.unresolved).toContain('non-inline-callback');
 const conditional=inventoryTestSource('conditional.test.ts',`import{test}from'bun:test';test.skipIf(flag)('conditional',()=>{});`)[0]!;expect(conditional.parameters).toBeUndefined();expect(conditional.reviewReasons).toContain('conditional-registration');
});
test('native inventory retains todo declarations without pretending they execute',()=>{
 const rows=inventoryTestSource('x.test.ts',`import{test}from'bun:test';test.todo('later');test.skip('skip',()=>{});test('now',()=>{});`);
 expect(rows.map(r=>r.lifecycle)).toEqual(['todo','skip','normal']);expect(rows[0]!.body).toBe('');expect(rows[0]!.unresolved).toEqual([]);
});
test('native inventory records body loops and only direct matcher expressions, not helper execution',()=>{
 const rows=inventoryTestSource('x.test.ts',`import{test,expect}from'bun:test';test('one',()=>{for(let i=0;i<2;i++)expect(i).toBeLessThan(2);helper();expect.objectContaining({x:1});});`);
 expect(rows[0]!.loops).toHaveLength(1);expect(rows[0]!.reviewReasons).toContain('body-loop');expect(rows[0]!.assertions).toEqual(['expect(i).toBeLessThan(2)']);expect(rows[0]!.calls).toContain('helper');
});
test('native inventory has deterministic hashes and refuses duplicate identities',async()=>{
 const root=await mkdtemp(join(tmpdir(),'bun-inventory-'));try{await mkdir(join(root,'tests/unit'),{recursive:true});await Bun.write(join(root,'tests/unit/a.test.ts'),`import{test}from'bun:test';test('same',()=>{});test('same',()=>{});`);await expect(inventoryNativeTests(root)).rejects.toThrow('Duplicate native test identities');}finally{await rm(root,{recursive:true,force:true});}
});
test('inventory drift check detects added tests, edited helpers and no-test file changes without rewriting report',async()=>{
 const root=await mkdtemp(join(tmpdir(),'bun-inventory-drift-'));try{
  await mkdir(join(root,'tests/unit'),{recursive:true});const file=join(root,'tests/unit/a.test.ts'),empty=join(root,'tests/unit/helper.test.ts'),report=join(root,'inventory.json');
  const source=`import{test}from'bun:test';const helper=1;test('one',()=>{});`;await Bun.write(file,source);await Bun.write(empty,'export const x=1;');const inv=await inventoryNativeTests(root);await Bun.write(report,JSON.stringify(inv,null,2)+'\n');const baseline=await Bun.file(report).text();await checkInventory(root,report);
  for(const change of [source.replace('helper=1','helper=2'),source+`test('two',()=>{});`]){await Bun.write(file,change);await expect(checkInventory(root,report)).rejects.toThrow('inventory stale');expect(await Bun.file(report).text()).toBe(baseline);}await Bun.write(file,source);await Bun.write(empty,'export const x=2;');await expect(checkInventory(root,report)).rejects.toThrow('inventory stale');
 }finally{await rm(root,{recursive:true,force:true});}
});
test('drift checks refuse unresolved callbacks even when a matching report exists',async()=>{
 const root=await mkdtemp(join(tmpdir(),'bun-inventory-opaque-'));try{await mkdir(join(root,'tests/unit'),{recursive:true});await Bun.write(join(root,'tests/unit/a.test.ts'),`import{test}from'bun:test';test('opaque',handler);`);const report=join(root,'inventory.json');await Bun.write(report,JSON.stringify(await inventoryNativeTests(root),null,2)+'\n');await expect(checkInventory(root,report)).rejects.toThrow('Unresolved native registrations');}finally{await rm(root,{recursive:true,force:true});}
});

import {reconcile} from '../../scripts/mapping-reconciliation.ts';
test('staging reconciliation distinguishes surviving IDs, missing IDs and absent mappings without credit',()=>{
 const cases=inventoryTestSource('x.test.ts',`import{test}from'bun:test';test('one',()=>{});test('two',()=>{});`),file={path:'staging.json',source:JSON.stringify({reviewState:'candidate-needs-parent-review',mappings:[{testId:cases[0]!.id},{testId:'bun:x.test.ts:old'}]})};
 const result=reconcile(cases,[file]);expect(result.identitiesPresent).toBe(1);expect(result.identitiesMissing).toBe(1);expect(result.withoutStagingCount).toBe(1);expect(result.mappings.every(m=>m.executionCredit===false&&m.sourceEquivalence==='unverified-staging-has-no-source-pin')).toBe(true);expect(()=>reconcile(cases,[file,file])).toThrow('duplicate staging identity');
});
test('inventory fails closed on imported binding shadowing, indirect aliases and tagged tables',async()=>{
 const root=await mkdtemp(join(tmpdir(),'bun-inventory-alias-'));try{await mkdir(join(root,'tests/unit'),{recursive:true});const file=join(root,'tests/unit/a.test.ts');
 for(const body of [`const t=test;t('missed',()=>{});`,`function f(test:any){test('shadow',()=>{})}`,"test.each`a | b`('tagged',()=>{})"]){await Bun.write(file,`import{test}from'bun:test';${body}`);const inventory=await inventoryNativeTests(root);expect(inventory.unresolved.length).toBeGreaterThan(0);}
 }finally{await rm(root,{recursive:true,force:true});}
});

test('native registration shadows in destructuring and deferred helper registration are unresolved',async()=>{
 const root=await mkdtemp(join(tmpdir(),'bun-inventory-scope-'));try{await mkdir(join(root,'tests/unit'),{recursive:true});const file=join(root,'tests/unit/a.test.ts');
 for(const body of [`function f({test}:any){test('shadow',()=>{})}`,`function register(){test('deferred',()=>{})}`,`const {it:another}=bt;another('indirect',()=>{});`]){await Bun.write(file,`import{test}from'bun:test';import * as bt from 'bun:test';${body}`);expect((await inventoryNativeTests(root)).unresolved.length).toBeGreaterThan(0);}
 }finally{await rm(root,{recursive:true,force:true});}
});

test('transparent TypeScript expressions cannot hide imported registration aliases',async()=>{
 const root=await mkdtemp(join(tmpdir(),'bun-inventory-ts-'));try{await mkdir(join(root,'tests/unit'),{recursive:true});
 for(const value of ['test as typeof test','test!','(test satisfies unknown)','<typeof test>test']){await Bun.write(join(root,'tests/unit/a.test.ts'),`import{test}from'bun:test';const spec=${value};spec('hidden',()=>{});`);expect((await inventoryNativeTests(root)).unresolved.length).toBeGreaterThan(0);}
 }finally{await rm(root,{recursive:true,force:true});}
});
test('nested function matchers are deferred syntax rather than direct test assertions',()=>{
 const rows=inventoryTestSource('x.test.ts',`import{test,expect}from'bun:test';test('one',()=>{const unused=()=>expect(1).toBe(2);});`);
 expect(rows[0]!.assertions).toEqual([]);expect(rows[0]!.deferredAssertions).toEqual(['expect(1).toBe(2)']);expect(rows[0]!.reviewReasons).toContain('deferred-assertion');expect(rows[0]!.reviewReasons).toContain('helper-or-no-direct-assertion');
});
test('test-support file edits invalidate the inventory independently of declaration files',async()=>{
 const root=await mkdtemp(join(tmpdir(),'bun-inventory-helper-'));try{await mkdir(join(root,'tests/unit'),{recursive:true});await mkdir(join(root,'tests/helpers'),{recursive:true});await Bun.write(join(root,'tests/unit/a.test.ts'),`import{test}from'bun:test';import{helper}from'../helpers/value.ts';test('one',()=>helper());`);const helper=join(root,'tests/helpers/value.ts');await Bun.write(helper,'export const helper=()=>1;');const report=join(root,'inventory.json');await Bun.write(report,JSON.stringify(await inventoryNativeTests(root),null,2)+'\n');await checkInventory(root,report);await Bun.write(helper,'export const helper=()=>2;');await expect(checkInventory(root,report)).rejects.toThrow('inventory stale');}finally{await rm(root,{recursive:true,force:true});}
});
