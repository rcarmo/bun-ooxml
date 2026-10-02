import {test,expect} from 'bun:test';
import {analyzeStaticReferences,parseStaticRange,remapStaticReferences,UniformXmlSnapshot,UniformApiError,uniformApiResult,analyzeFormulaReferences,type UniformApiCategory} from '../../src/index.ts';
const refusal=(fn:()=>unknown,category:UniformApiCategory)=>{const r=uniformApiResult(fn);expect(r).toEqual({ok:false,category,value:null});};
test('uniform formula enforces arities while legacy remains unchanged and records are detached',()=>{
 for(const source of ['IF(A1)','IF(A1,B2,C3,D4)','ABS(A1,B2)','ROUND(A1)','SUM()','NOT(A1,B2)'])refusal(()=>analyzeStaticReferences(source),'unsupported-static-reference');
 expect(analyzeFormulaReferences('ABS(A1,B2)')).toHaveLength(2);
 expect(analyzeStaticReferences('IF(A1,B2)')).toHaveLength(2);expect(analyzeStaticReferences('round(A1,2)')).toHaveLength(1);
 const s='IF(A1="B2",\'O\'\'Brien\'!$C$4,SUM(D1:E2))',refs=analyzeStaticReferences(s);
 expect(refs.map(r=>[r.start,r.end])).toEqual([[3,5],[11,26],[31,36]]);expect(refs[0]!.last).toEqual(refs[0]!.first);expect(refs[0]!.last).not.toBe(refs[0]!.first);
 refs[0]!.first.row=99;expect(analyzeStaticReferences(s)[0]!.first.row).toBe(1);
 refusal(()=>analyzeStaticReferences('x'.repeat(1048577)),'static-reference-limit');
 refusal(()=>analyzeStaticReferences('('.repeat(130)+'A1'+')'.repeat(130)),'static-reference-limit');
});
test('uniform range/remap endpoints, categories and caller custody',()=>{
 expect(parseStaticRange('$3:$1')).toEqual({sheet:'',axis:'row',first:{row:3,column:0,rowAbsolute:true,columnAbsolute:false},last:{row:1,column:0,rowAbsolute:true,columnAbsolute:false}});
 refusal(()=>parseStaticRange('A1#'),'unsupported-direct-range');
 const edit=Object.freeze({axis:'row' as const,at:2,count:1,sheet:'Main'}),s='a1+$B$2+"C3"';
 expect(remapStaticReferences(s,'Main',edit)).toBe('a1+$B$3+"C3"');expect(edit).toEqual({axis:'row',at:2,count:1,sheet:'Main'});
 refusal(()=>remapStaticReferences('A1048576','Main',edit),'invalid-reference-insertion');
 refusal(()=>remapStaticReferences('ABS(A1,B2)','Main',edit),'unsupported-static-reference');
});
test('uniform XML issued UTF16 targets, typed refusals and reusable atomic custody',()=>{
 const source='<r>😀<a x=\'one\'/><b>keep</b></r>',s=UniformXmlSnapshot.parse(source),a=s.elements[1]!,foreign=UniformXmlSnapshot.parse(source).elements[1]!;
 expect([a.start,a.end]).toEqual([5,17]);expect(source.slice(a.start,a.end)).toBe("<a x='one'/>");
 const patches=Object.freeze([Object.freeze({target:a,name:'x',value:'two'})]);
 expect(s.setAttributes(patches)).toBe("<r>😀<a x='two'/><b>keep</b></r>");expect(patches[0]!.value).toBe('two');expect(s.remove([])).toBe(source);
 refusal(()=>s.remove([s.elements[0]!]),'root');refusal(()=>s.remove([a,a]),'overlap-or-duplicate');refusal(()=>s.remove([foreign]),'foreign-target');
 refusal(()=>s.setAttributes([{target:a,name:'xmlns',value:'x'}]),'invalid-name-or-value');
 refusal(()=>s.setAttributes([{target:a,name:'x',value:'\ud800'}]),'unsafe-XML');
 refusal(()=>UniformXmlSnapshot.parse('\ufeff<r/>'),'invalid-name-or-value');refusal(()=>UniformXmlSnapshot.parse('<r>\ud800</r>'),'unsafe-XML');
 refusal(()=>UniformXmlSnapshot.parse('<r>'+'x'.repeat(8388608)+'</r>'),'xml-edit-limit');
 refusal(()=>s.remove(new Array(100001).fill(a)),'xml-edit-limit');
 expect(s.remove([a])).toBe('<r>😀<b>keep</b></r>');expect(s.setAttributes([])).toBe(source);
});
test('uniform XML deterministic insertion readback and expanded attribute duplicate refusal',()=>{
 const source='<root xmlns="u" xmlns:n1="occupied"><a/><b>keep</b></root>',s=UniformXmlSnapshot.parse(source),a=s.elements[1]!;
 const children=[{name:{namespaceURI:'new',localName:'x'},attributes:[{name:{namespaceURI:'other',localName:'a'},value:'value'}],children:[{name:{namespaceURI:'',localName:'plain'},children:['text']}]}];
 expect(s.appendChildren([{target:a,children}])).toBe('<root xmlns="u" xmlns:n1="occupied"><a><n2:x xmlns:n2="new" xmlns:n3="other" n3:a="value"><plain xmlns="">text</plain></n2:x></a><b>keep</b></root>');
 refusal(()=>s.appendChildren([{target:a,children:[{name:{namespaceURI:'',localName:'x'},attributes:[{name:{namespaceURI:'',localName:'a'},value:'x'},{name:{namespaceURI:'',localName:'a'},value:'y'}]}]}]),'overlap-or-duplicate');
 expect(s.appendChildren([])).toBe(source);
});
test('uniform sealed lexical and remap boundaries discriminate unsupported syntax and output growth',()=>{
 for(const source of ['1.5','1e2','IF(A1,B2)','A1 + B2',"'α sheet'!A1"])expect(analyzeStaticReferences(source)).toBeArray();
 for(const source of ['A١','A1\t+B2','A1\u00a0+B2','A1 : B2',"'Main' !A1",'SUM("\u0001",A1)'])refusal(()=>analyzeStaticReferences(source),'unsupported-static-reference');
 for(const source of [' A1','A1 ','A1 : B2','A١',"'x\u0080'!A1","'x\u009f'!A1",'A1'.repeat(524289)])refusal(()=>parseStaticRange(source),'unsupported-direct-range');
 for(const sheet of ['', ' ', 'Main\t', 'A:B', '\ud800','x'.repeat(256)])refusal(()=>remapStaticReferences('A1','Main',{axis:'row',at:1,count:1,sheet}),'invalid-reference-insertion');
 expect(remapStaticReferences('a1:b2','Main',{axis:'row',at:2,count:1,sheet:'Main'})).toBe('A1:B3');
 expect(remapStaticReferences('A1','ΟΣ',{axis:'row',at:1,count:1,sheet:'οσ'})).toBe('A2');
 expect(remapStaticReferences('A1','İ',{axis:'row',at:1,count:1,sheet:'i'})).toBe('A2');
 expect(remapStaticReferences('A1','ß',{axis:'row',at:1,count:1,sheet:'ss'})).toBe('A1');
 const growth='A9&"'+'x'.repeat(1048571)+'"';expect(new TextEncoder().encode(growth)).toHaveLength(1048576);
 refusal(()=>remapStaticReferences(growth,'Main',{axis:'row',at:1,count:1,sheet:'Main'}),'static-reference-limit');
});
test('uniform formula exact source token depth reference and arity budget boundaries',()=>{
 const bytes='"'+'x'.repeat(1048574)+'"';expect(analyzeStaticReferences(bytes)).toEqual([]);refusal(()=>analyzeStaticReferences(bytes+' '),'static-reference-limit');
 const utf8='"'+'é'.repeat(524287)+'"';expect(analyzeStaticReferences(utf8)).toEqual([]);refusal(()=>analyzeStaticReferences(utf8+' '),'static-reference-limit');
 const tokens='1+'.repeat(49999)+'1%';expect(analyzeStaticReferences(tokens)).toEqual([]);refusal(()=>analyzeStaticReferences(tokens+'%'),'static-reference-limit');
 expect(analyzeStaticReferences('('.repeat(128)+'A1'+')'.repeat(128))).toHaveLength(1);refusal(()=>analyzeStaticReferences('('.repeat(129)+'A1'+')'.repeat(129)),'static-reference-limit');
 expect(analyzeStaticReferences('A1+'.repeat(9999)+'A1')).toHaveLength(10000);refusal(()=>analyzeStaticReferences('A1+'.repeat(10000)+'A1'),'static-reference-limit');
 for(const name of ['SUM','MIN','MAX','AVERAGE','COUNT','COUNTA','AND','OR']){expect(analyzeStaticReferences(name+'('+new Array(255).fill('1').join(',')+')')).toEqual([]);refusal(()=>analyzeStaticReferences(name+'('+new Array(256).fill('1').join(',')+')'),'unsupported-static-reference');}
});
test('uniform XML source element depth and output budgets are independently typed',()=>{
 const exact='<r>'+'x'.repeat(8388601)+'</r>';expect(UniformXmlSnapshot.parse(exact).remove([])).toBe(exact);refusal(()=>UniformXmlSnapshot.parse(exact+' '),'xml-edit-limit');
 const nodes='<r>'+'<a/>'.repeat(99999)+'</r>';expect(UniformXmlSnapshot.parse(nodes).elements).toHaveLength(100000);refusal(()=>UniformXmlSnapshot.parse(nodes.replace('</r>','<a/></r>')),'xml-edit-limit');
 const depth='<r>'.repeat(256)+'</r>'.repeat(256);expect(UniformXmlSnapshot.parse(depth).elements).toHaveLength(256);refusal(()=>UniformXmlSnapshot.parse('<r>'+depth+'</r>'),'xml-edit-limit');
 const s=UniformXmlSnapshot.parse('<r><a/></r>'),a=s.elements[1]!;
 refusal(()=>s.setAttributes([{target:a,name:'x',value:'x'.repeat(8388608)}]),'xml-edit-limit');
 refusal(()=>s.appendChildren([{target:a,children:['x'.repeat(8388608)]}]),'xml-edit-limit');
 const large=UniformXmlSnapshot.parse('<r>'+'<a/>'.repeat(99998)+'</r>'),child={name:{localName:'n',namespaceURI:''}};
 refusal(()=>large.appendChildren([{target:large.elements[0]!,children:[child,child]}]),'xml-edit-limit');
 expect(large.appendChildren([])).toBe('<r>'+'<a/>'.repeat(99998)+'</r>');expect(s.remove([])).toBe('<r><a/></r>');
});
test('uniform XML no-op and refusal preserve frozen caller batches and ignore array iterators',()=>{
 const s=UniformXmlSnapshot.parse('<r>]]<a/>>😀<b/></r>'),a=s.elements[1]!,b=s.elements[2]!,foreign=UniformXmlSnapshot.parse('<r><a/></r>').elements[1]!;
 refusal(()=>s.remove([a]),'unsafe-XML');refusal(()=>s.replaceElements([{target:s.elements[0]!,children:[]}]),'root');
 refusal(()=>s.appendChildren([{target:s.elements[0]!,children:[]},{target:a,children:[]}]),'overlap-or-duplicate');
 refusal(()=>s.setAttributes([{target:foreign,name:'x',value:'x'}]),'foreign-target');refusal(()=>s.replaceElements([{target:foreign,children:[]}]),'foreign-target');
 const batch=[b];Object.defineProperty(batch,Symbol.iterator,{value:()=>{throw new Error('untrusted iterator');}});Object.freeze(batch);expect(s.remove(batch)).toBe('<r>]]<a/>>😀</r>');
 expect(s.appendChildren([])).toBe('<r>]]<a/>>😀<b/></r>');
});
test('adapter catches typed profile refusal only',()=>{expect(()=>uniformApiResult(()=>{throw new Error('programmer fault');})).toThrow('programmer fault');expect(new UniformApiError('root','root')).toBeInstanceOf(Error);});
