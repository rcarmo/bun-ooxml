import {expect,test} from 'bun:test';
import {mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {Presentation} from '../../src/pptx/index.ts';
import {pictureRecipes,pictureInput,pictureContract} from '../helpers/picture-inputs.ts';
import {archive} from '../helpers/contract20-oracle.ts';
function membersSame(a:Uint8Array,b:Uint8Array){
 const left=archive(a),right=archive(b);
 expect([...right.keys()].sort()).toEqual([...left.keys()].sort());
 for(const[n,bytes]of left)expect(right.get(n)).toEqual(bytes);
}
test('shared picture binding has exactly six IDs and twelve cases, without execution credit',()=>{
 expect(pictureRecipes).toHaveLength(pictureContract.counts.cases);
 expect(new Set(pictureRecipes.map(c=>c.scenarioId)).size).toBe(pictureContract.counts.ids);
 expect(pictureContract.executionCredit).toBe(false);
});
for(const recipe of pictureRecipes)test(`${recipe.scenarioId} [${recipe.caseId}]`,async()=>{
 const bytes=await pictureInput(recipe),before=bytes.slice(),p=await Presentation.open(bytes);
 const originalFetch=globalThis.fetch;
 globalThis.fetch=(()=>{throw Error('picture inspection fetched an external asset');}) as unknown as typeof fetch;
 try{
  if(recipe.errorCode){
   expect(()=>p.slides[0]!.inspectPictures()).toThrow(expect.objectContaining({code:recipe.errorCode}));
   expect(()=>p.slides[0]!.inspectPictures()).toThrow(expect.objectContaining({code:recipe.errorCode}));
  }else{
   if(!recipe.expected)throw Error('Shared success recipe lacks expected records');
   const rows=p.slides[0]!.inspectPictures();expect(rows).toEqual(recipe.expected);
   // Mutate every exposed nested record; later reads must remain detached.
   for(const row of rows){
    row.shapeId=-1;row.name='changed';row.description='changed';row.slidePart='changed';
    if(row.embedded){row.embedded.target='changed';row.embedded.byteLength=-1;}
    if(row.linked)row.linked.target='changed';
    if(row.transform)row.transform.x=-999;
    row.crop.left=-999;
    for(const group of row.groups){group.name='changed';if(group.transform)group.transform.childX=-999;}
    row.groups.length=0;
   }
   expect(p.slides[0]!.inspectPictures()).toEqual(recipe.expected);
   const dir=await mkdtemp(join(tmpdir(),'pptx-picture-inspection-'));
   try{
    const path=join(dir,'copy.pptx');await p.save(path);const saved=await Bun.file(path).bytes();
    membersSame(bytes,saved);expect((await Presentation.open(saved)).slides[0]!.inspectPictures()).toEqual(recipe.expected);
   }finally{await rm(dir,{recursive:true,force:true});}
  }
  expect(bytes).toEqual(before);expect(p.package.diff()).toEqual({added:[],changed:[],removed:[]});membersSame(bytes,p.package.toBytes());
 }finally{globalThis.fetch=originalFetch;}
});
test('native picture API returns an empty collection for a newly created text slide',()=>{
 const p=Presentation.create();expect(p.addTextSlide('No pictures').inspectPictures()).toEqual([]);
});
