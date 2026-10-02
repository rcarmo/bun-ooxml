import {test,expect,spyOn} from 'bun:test';
import * as admission from '../../src/opc/admission.ts';
import * as visibility from '../../src/pptx/visibility.ts';
import {OpcPackage} from '../../src/opc/index.ts';
import {fixturePath} from '../../scripts/fixture-inputs.ts';
import {parseXml,attribute} from '../../src/xml/index.ts';
const valid='fixture-d9d6a313182a71a73d75a26a0ff3b7826dbd2e300e1d202114ec9f8fb018fda5';
const retained='fixture-e01ded1106a28f94a3439e8368f9a12ec360891f4a9e2810f6504c4c328ed79c',control='fixture-fa245a3df00fef7f7bf4739921ee840194040161e06490589e3d52cc9fa7a71d';
test('invalid admission configuration refuses before obtaining any filesystem source',async()=>{
 const original=await Bun.file(fixturePath(valid)).bytes();
 for(const budget of ['maxArchiveBytes','maxEntries']as const){
  const probes={file:spyOn(Bun,'file').mockImplementation(()=>{throw Error('source accessed before argument refusal');})};
  try{await expect(admission.admitPackageFile(fixturePath(valid),{[budget]:-1})).rejects.toMatchObject({code:'package-admission-limit-invalid'});expect(probes.file).not.toHaveBeenCalled();}
  finally{probes.file.mockRestore();}
  const parts=await admission.admitPackageFile(fixturePath(valid));expect(parts.has('word/document.xml')).toBe(true);
  await expect(admission.admitPackageFile(fixturePath(valid),{[budget]:0})).rejects.toMatchObject({code:budget==='maxArchiveBytes'?'zip-archive-too-large':'zip-too-many-entries'});
  expect(await Bun.file(fixturePath(valid)).bytes()).toEqual(original);
 }
});
test('retained namespaced marker is observed independently of unqualified show and rendering',async()=>{
 for(const [id,source]of [[retained,true],[control,false]]as const){
  const caller=await Bun.file(fixturePath(id)).bytes(),before=caller.slice(),pkg=await OpcPackage.open(caller),observations=visibility.inspectSlideVisibility(pkg);
  expect(observations.map(s=>[s.ordinal,s.slideId,s.relationshipId,s.part])).toEqual([1,2,3,4].map((n,i)=>[n,String(256+i),'rId'+(source?7+i:2+i),'ppt/slides/slide'+n+'.xml']));
  expect(observations.map(s=>s.show)).toEqual([null,null,null,null]);expect(observations.map(s=>s.namespacedShow)).toEqual([null,null,source?'0':null,null]);
  expect(observations.every(s=>s.declaredHidden===false&&s.applicationConfirmedHidden===false)).toBe(true);
  expect(caller).toEqual(before);expect(pkg.toBytes()).toEqual(before);expect(Object.isFrozen(observations)).toBe(true);expect(observations.every(Object.isFrozen)).toBe(true);
 }
});
test('visibility inspector follows relationship list order and namespace meaning without editing',async()=>{
 const pkg=await OpcPackage.open(await Bun.file(fixturePath(control)).bytes()),part='ppt/presentation.xml',xml=pkg.text(part),d=parseXml(xml),list=d.root.children.find(n=>n.localName==='sldIdLst')!,a=list.children[0]!,b=list.children[1]!;
 pkg.set(part,xml.slice(0,a.start)+xml.slice(b.start,b.end)+xml.slice(a.start,a.end)+xml.slice(b.end));
 const bytes=pkg.toBytes(),rows=visibility.inspectSlideVisibility(pkg);expect(rows.map(s=>s.slideId)).toEqual(['257','256','258','259']);expect(rows[0]!.part).toBe('ppt/slides/slide2.xml');expect(pkg.toBytes()).toEqual(bytes);
 const slide=pkg.text(rows[0]!.part);pkg.set(rows[0]!.part,slide.replace('<p:sld ','<p:sld show="0" '));expect(visibility.inspectSlideVisibility(pkg)[0]!.declaredHidden).toBe(true);
 const updated=pkg.text(rows[0]!.part);pkg.set(rows[0]!.part,updated.replace('show="0"','show="maybe"'));expect(()=>visibility.inspectSlideVisibility(pkg)).toThrow();
 expect(attribute(parseXml(updated).root,'show')).toBe('0');
});
