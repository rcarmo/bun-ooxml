import assert from 'node:assert/strict';
import type {StepBinding} from '../../scripts/gherkin.ts';
import {analyzeStaticReferences,parseStaticRange,remapStaticReferences,UniformXmlSnapshot,UniformApiError,type FormulaReference,type FormulaInsertion,type XmlContent,uniformApiResult} from '../../src/index.ts';
import {parseXml,attribute} from '../../src/xml/index.ts';
export const scenarioIds=['@id-xml-go-attribute-splice-custody','@id-xml-go-attribute-batch-refusal','@id-xml-go-child-insertion-custody','@id-xml-go-child-insertion-refusal','@id-xml-go-element-removal-custody','@id-xml-go-element-removal-refusal','@id-xml-go-element-replacement-custody','@id-xml-go-element-replacement-refusal','@id-xml-go-immutable-leaf-seed','@id-xlsx-go-formula-analysis-counts','@id-xlsx-go-formula-quoted-sheet-flags','@id-xlsx-go-formula-analysis-refusal','@id-xlsx-go-formula-literal-punctuation','@id-xlsx-go-direct-range-parsing','@id-xlsx-go-direct-range-refusal','@id-xlsx-go-static-remap-exact','@id-xlsx-go-static-remap-refusal','@id-xlsx-go-static-reference-properties'];
type Matrix={source:string;refs:FormulaReference[];expected:FormulaReference[];wrapped:FormulaReference[];remapped:string};
type State={source?:string;original?:string;snapshot?:UniformXmlSnapshot;output?:string;refs?:FormulaReference[];range?:ReturnType<typeof parseStaticRange>;error?:unknown;patchesBefore?:string;patches?:unknown;caller?:Uint8Array;callerBefore?:Uint8Array;matrix?:Matrix[];operation?:()=>string;targetsBefore?:string};
const state=(c:Record<string,unknown>)=>{c.uniform??={};return c.uniform as State;};
const id=(c:Record<string,unknown>)=>(c.scenario as {id:string}).id;
import profilePin from '../../docs/behaviors/uniform-api18-candidate.json';
const isProfile=(c:Record<string,unknown>)=>{const f=c.feature as {path:string;sourceSha256:string};const path=f.path.replace(/^references\/fixtures-ooxml\//,'');return scenarioIds.includes(id(c))&&profilePin.featureSeals[path as keyof typeof profilePin.featureSeals]===f.sourceSha256;};
const element=(localName:string,namespaceURI='',children:readonly XmlContent[]=[]):XmlContent=>({name:{localName,namespaceURI},children});
function target(s:State,local:string){const t=s.snapshot?.elements.find(t=>t.localName===local);assert(t);return t;}
function invoke(s:State,operation:()=>unknown){s.error=undefined;try{const v=operation();if(typeof v==='string')s.output=v;else if(Array.isArray(v))s.refs=v;else s.range=v as State['range'];}catch(error){s.error=error;}}
function xmlEdit(s:State,patches:unknown,fn:()=>string){s.operation=fn;s.patches=patches;s.patchesBefore=JSON.stringify(patches);invoke(s,fn);assert.equal(JSON.stringify(patches),s.patchesBefore);}
function refused(s:State,categories:string[]){assert(s.error instanceof UniformApiError);assert(categories.includes(s.error.category));assert.equal(s.output,undefined);assert.equal(s.refs,undefined);assert.equal(s.range,undefined);}
const selectedBindings:StepBinding[]=[
 {pattern:/^the XML source is (.+)$/,run:(c,source)=>{const s=state(c);s.source=s.original=source!;s.snapshot=UniformXmlSnapshot.parse(source!);s.targetsBefore=JSON.stringify(s.snapshot.elements);}},
 {pattern:/^the formula source is JSON (.+?)(?: in context sheet Main)?$/,run:(c,json)=>{const s=state(c);s.source=s.original=JSON.parse(json!);assert.equal(typeof s.source,'string');}},
 {pattern:/^the direct range source is JSON (.+)$/,run:(c,json)=>{const s=state(c);s.source=s.original=JSON.parse(json!);assert.equal(typeof s.source,'string');}},
 {pattern:/^the static formula analyser reads the source$/,run:c=>{const s=state(c);invoke(s,()=>analyzeStaticReferences(s.source!));}},
 {pattern:/^the direct-range parser reads the source$/,run:c=>{const s=state(c);invoke(s,()=>parseStaticRange(s.source!));}},
 {pattern:/^it returns (\d+) reference records without error$/,run:(c,n)=>{const s=state(c);assert.equal(s.error,undefined);assert.equal(s.refs?.length,Number(n));}},
 {pattern:/^every reference has a nonempty byte span inside the original source$/,run:c=>{const s=state(c);assert(s.refs);for(const r of s.refs){assert(r.start>=0&&r.end>r.start&&r.end<=new TextEncoder().encode(s.source!).length);}}},
 {pattern:/^the complete normalized reference records equal JSON (.+)$/,run:(c,json)=>{assert.deepEqual(state(c).refs??[],JSON.parse(json!));}},
 {pattern:/^the complete normalized reference records or refusal equal JSON (.+)$/,run:(c,json)=>{assert.deepEqual(state(c).refs??null,JSON.parse(json!));}},
 {pattern:/^the complete normalized direct range equals JSON (.+)$/,run:(c,json)=>{assert.deepEqual(state(c).range,JSON.parse(json!));}},
 {pattern:/^its one reference spans every byte of the original source$/,run:c=>{const s=state(c);assert.equal(s.refs?.length,1);assert.deepEqual([s.refs![0]!.start,s.refs![0]!.end],[0,new TextEncoder().encode(s.source!).length]);}},
 {pattern:/^the sheet is O'Brien with first cell column (\d+) row (\d+) and absolute column only$/,run:(c,col,row)=>{const r=state(c).refs?.[0];assert(r);assert.equal(r.sheet,"O'Brien");assert.equal(r.axis,'cell');assert.deepEqual(r.first,{column:Number(col),row:Number(row),columnAbsolute:true,rowAbsolute:false});}},
 {pattern:/^the last cell is column (\d+) row (\d+) with absolute row only$/,run:(c,col,row)=>{assert.deepEqual(state(c).refs?.[0]?.last,{column:Number(col),row:Number(row),columnAbsolute:false,rowAbsolute:true});}},
 {pattern:/^it returns an error and zero reference records$/,run:c=>refused(state(c),['unsupported-static-reference'])},
 {pattern:/^the analysis is (accepted without error|refused with zero references)$/,run:(c,result)=>{const s=state(c);if(result==='accepted without error'){assert.equal(s.error,undefined);assert(Array.isArray(s.refs));}else refused(s,['unsupported-static-reference']);}},
 {pattern:/^its sheet is JSON (.+) and its axis is (cell|row|column)$/,run:(c,sheet,axis)=>{const s=state(c);assert.equal(s.error,undefined);assert.equal(s.range?.sheet,JSON.parse(sheet!));assert.equal(s.range?.axis,axis);}},
 {pattern:/^its first coordinate has row (\d+) and column (\d+)$/,run:(c,row,col)=>{assert.equal(state(c).range?.first.row,Number(row));assert.equal(state(c).range?.first.column,Number(col));}},
 {pattern:/^it returns an error$/,run:c=>refused(state(c),['unsupported-direct-range'])},
 {pattern:/^the static remapper inserts (\w+) at (\d+) by (\d+) on sheet JSON (.+)$/,run:(c,axis,at,count,sheet)=>{const s=state(c),change={axis:axis as FormulaInsertion['axis'],at:Number(at),count:Number(count),sheet:JSON.parse(sheet!)};xmlEdit(s,change,()=>remapStaticReferences(s.source!,'Main',change));}},
 {pattern:/^the complete replacement expression equals JSON (.+)$/,run:(c,json)=>{assert.equal(state(c).error,undefined);assert.equal(state(c).output,JSON.parse(json!));}},
 {pattern:/^it returns an error and an empty replacement expression$/,run:c=>refused(state(c),['invalid-reference-insertion','unsupported-static-reference'])},
 {pattern:/^cell tokens A1, \$B2, C\$3 and \$XFD\$9$/,run:()=>{}},
 {pattern:/^optional prefixes empty, Main! and 'Input Data'! with operators \+, -, \*, \/, & and >=$/,run:()=>{}},
 {pattern:/^the static analyser checks all 3 by 4 by 4 by 6 source expressions$/,run:c=>{
  const s=state(c);s.matrix=[];
  // Independent finite token goldens, never produced by a parser under test.
  const coordinates=[{row:1,column:1,rowAbsolute:false,columnAbsolute:false},{row:2,column:2,rowAbsolute:false,columnAbsolute:true},{row:3,column:3,rowAbsolute:true,columnAbsolute:false},{row:9,column:16384,rowAbsolute:true,columnAbsolute:true}],tokens=['A1','$B2','C$3','$XFD$9'];
  for(const prefix of ['', 'Main!',"'Input Data'!"])for(let a=0;a<4;a++)for(let b=0;b<4;b++)for(const op of ['+','-','*','/','&','>=']){
   const first=prefix+tokens[a],second=prefix+tokens[b],source=first+op+second,sheet=prefix===''?'':prefix==='Main!'?'Main':'Input Data';
   const record=(i:number,start:number,end:number):FormulaReference=>({sheet,axis:'cell',first:{...coordinates[i]!},last:{...coordinates[i]!},start,end});
   const expected=[record(a,0,first.length),record(b,first.length+op.length,source.length)];
   s.matrix.push({source,expected,refs:analyzeStaticReferences(source),wrapped:analyzeStaticReferences('SUM('+source+')'),remapped:remapStaticReferences(source,'Main',{axis:'row',at:100,count:1,sheet:'Main'})});
  }
 }},
 {pattern:/^every expression has two references whose source slices each parse as one matching reference$/,run:c=>{const rows=state(c).matrix;assert.equal(rows?.length,288);for(const row of rows!){assert.deepEqual(row.refs,row.expected);const bytes=new TextEncoder().encode(row.source);for(const r of row.expected){const slice=new TextDecoder().decode(bytes.slice(r.start,r.end));assert.deepEqual(analyzeStaticReferences(slice),[{...r,start:0,end:r.end-r.start}]);}}}},
 {pattern:/^inserting one row at 100 on Main leaves each original expression byte-identical$/,run:c=>{assert.equal(state(c).matrix?.length,288);for(const row of state(c).matrix!)assert.equal(row.remapped,row.source);}},
 {pattern:/^wrapping each expression in SUM preserves its references after adjusting their byte spans$/,run:c=>{assert.equal(state(c).matrix?.length,288);for(const row of state(c).matrix!)assert.deepEqual(row.wrapped,row.expected.map(r=>({...r,start:r.start+4,end:r.end+4})));}},
 {pattern:/^a lexical edit sets the attribute (\S+) of the first t element to (.+)$/,run:(c,name,value)=>{const s=state(c),patches=[{target:target(s,'t'),name:name!,value:value!}];xmlEdit(s,patches,()=>s.snapshot!.setAttributes(patches));}},
 {pattern:/^one lexical edit batch sets a of the first t element to x and to y$/,run:c=>{const s=state(c),patches=['x','y'].map(value=>({target:target(s,'t'),name:'a',value}));xmlEdit(s,patches,()=>s.snapshot!.setAttributes(patches));}},
 {pattern:/^the edit returns an error instead of accepting that batch$/,run:c=>refused(state(c),['overlap-or-duplicate'])},
 {pattern:/^the complete output bytes equal (.+)$/,run:(c,out)=>{assert.equal(state(c).output,out);}},
 {pattern:/^a structured edit inserts a new-namespace x child with an other-namespace a attribute and a plain text child under the existing a element$/,run:c=>{const s=state(c),patches=[{target:target(s,'a'),children:[{name:{localName:'x',namespaceURI:'new'},attributes:[{name:{localName:'a',namespaceURI:'other'},value:'value'}],children:[element('plain','',['text'])]}]}];xmlEdit(s,patches,()=>s.snapshot!.appendChildren(patches));}},
 {pattern:/^reparsing finds expanded element names new\/x and empty-namespace plain$/,run:c=>{const x=parseXml(state(c).output!).root.children[0]!.children[0]!;assert.deepEqual([x.namespaceURI,x.localName],['new','x']);assert.deepEqual([x.children[0]!.namespaceURI,x.children[0]!.localName],['','plain']);}},
 {pattern:/^the authored attribute expanded name is other\/a with JSON value (.+) and the plain grandchild text is JSON (.+)$/,run:(c,value,text)=>{const x=parseXml(state(c).output!).root.children[0]!.children[0]!;assert.equal(attribute(x,'a','other'),JSON.parse(value!));assert.equal(x.children[0]!.text,JSON.parse(text!));}},
 {pattern:/^the unedited sibling bytes (.+) remain in the output$/,run:(c,sibling)=>{assert(state(c).source?.includes(sibling!));assert(state(c).output?.includes(sibling!));}},
 {pattern:/^one structured insertion batch targets both the root and its nested a element$/,run:c=>{const s=state(c),patches=[s.snapshot!.elements[0]!,target(s,'a')].map(target=>({target,children:[element('x')]}));xmlEdit(s,patches,()=>s.snapshot!.appendChildren(patches));}},
 {pattern:/^the insertion returns an error$/,run:c=>refused(state(c),['overlap-or-duplicate'])},
 {pattern:/^a separate empty (insertion batch|removal|replacement) returns the exact original source bytes$/,run:c=>{const s=state(c);assert.equal(s.snapshot!.remove([]),s.source);assert.equal(s.snapshot!.appendChildren([]),s.source);assert.equal(s.snapshot!.replaceElements([]),s.source);}},
 {pattern:/^the XML editor removes the p:a subtree and the p:c element from one parsed snapshot$/,run:c=>{const s=state(c),patches=[target(s,'a'),target(s,'c')];xmlEdit(s,patches,()=>s.snapshot!.remove(patches));}},
 {pattern:/^a removal batch selects (root|p:a and its nested p:b)$/,run:(c,selection)=>{const s=state(c),patches=selection==='root'?[s.snapshot!.elements[0]!]:[target(s,'a'),target(s,'b')];xmlEdit(s,patches,()=>s.snapshot!.remove(patches));}},
 {pattern:/^the removal returns an error$/,run:c=>refused(state(c),['root','overlap-or-duplicate'])},
 {pattern:/^the XML editor replaces p:old with a bound-namespace new element containing value and an empty-namespace plain element$/,run:c=>{const s=state(c),patches=[{target:target(s,'old'),children:[element('new','bound',['value']),element('plain')]}];xmlEdit(s,patches,()=>s.snapshot!.replaceElements(patches));}},
 {pattern:/^a replacement batch selects (root|p:old twice|p:old and its nested p:child)$/,run:(c,selection)=>{const s=state(c),targets=selection==='root'?[s.snapshot!.elements[0]!]:selection==='p:old twice'?[target(s,'old'),target(s,'old')]:[target(s,'old'),target(s,'child')],patches=targets.map(target=>({target,children:[element('new')]}));xmlEdit(s,patches,()=>s.snapshot!.replaceElements(patches));}},
 {pattern:/^replacement returns an error and no edited output$/,run:c=>refused(state(c),['root','overlap-or-duplicate'])},
 {pattern:/^the XML editor parses a caller-owned byte slice and performs an empty edit$/,run:c=>{const s=state(c);s.caller=new TextEncoder().encode(s.source!);s.callerBefore=s.caller.slice();s.snapshot=UniformXmlSnapshot.parse(new TextDecoder('utf-8',{fatal:true}).decode(s.caller));s.targetsBefore=JSON.stringify(s.snapshot.elements);s.output=s.snapshot.remove([]);}},
 {pattern:/^the caller input bytes still equal the original XML source$/,run:c=>{const s=state(c);assert.deepEqual(s.caller,s.callerBefore);assert.equal(new TextDecoder().decode(s.caller),s.source);}},
 {pattern:/^the empty edit returns the exact original source bytes$/,run:c=>assert.equal(state(c).output,state(c).source)},
 {pattern:/^the uniform profile result, refusal category and immutable input custody match the sealed API contract$/,run:c=>{
  const s=state(c);assert.equal(s.source,s.original);if(s.patchesBefore!==undefined)assert.equal(JSON.stringify(s.patches),s.patchesBefore);
  if(s.error){assert(s.error instanceof UniformApiError);
   const scenario=id(c),values=(c.case as {example?:{values:Record<string,string>}}).example?.values??{};
   const expected=scenario.includes('formula-')?'unsupported-static-reference':scenario.includes('direct-range-')?'unsupported-direct-range':scenario.includes('static-remap-')?(values.variant==='dynamic reference'?'unsupported-static-reference':'invalid-reference-insertion'):values.selection==='root'?'root':'overlap-or-duplicate';
   assert.equal(s.error.category,expected);assert.deepEqual(uniformApiResult(()=>{throw s.error;}),{ok:false,category:expected,value:null});
  }
  if(s.operation){const replay=uniformApiResult(s.operation);if(s.error){assert.equal(replay.ok,false);if(!replay.ok)assert.equal(replay.category,(s.error as UniformApiError).category);}else assert.deepEqual(replay,{ok:true,value:s.output});}
  if(s.snapshot){assert.equal(JSON.stringify(s.snapshot.elements),s.targetsBefore);assert.equal(s.snapshot.remove([]),s.original);}
  if(!s.error&&id(c).startsWith('@id-xml-')){assert.equal(s.error,undefined);assert(s.snapshot);assert.equal(s.snapshot.remove([]),s.original);assert.equal(s.snapshot.setAttributes([]),s.original);assert.equal(s.snapshot.appendChildren([]),s.original);assert.equal(s.snapshot.replaceElements([]),s.original);for(const t of s.snapshot.elements){assert(Object.isFrozen(t));assert(s.original!.slice(t.start,t.end).startsWith('<'));}}
  else if(s.matrix){assert.equal(s.matrix.length,288);for(const row of s.matrix)assert.deepEqual(analyzeStaticReferences(row.source),row.expected);}
  else if(s.refs){const fresh=analyzeStaticReferences(s.source!);assert.deepEqual(fresh,s.refs);if(s.refs[0]){const before=structuredClone(fresh);s.refs[0].first.row++;assert.deepEqual(analyzeStaticReferences(s.source!),before);s.refs=before;}}
  else if(s.range){const before=structuredClone(s.range);s.range.first.row++;assert.deepEqual(parseStaticRange(s.source!),before);s.range=before;}
 }},
];
export function withUniformApi18(bindings:readonly StepBinding[]):StepBinding[]{
 const extras=selectedBindings.filter(b=>b.pattern.source.startsWith('^the complete normalized')||b.pattern.source.startsWith('^the authored attribute')||b.pattern.source.startsWith('^the uniform profile'));
 return [...bindings.map(b=>({...b,run:(c:Record<string,unknown>,...args:string[])=>{
  if(!isProfile(c))return b.run(c,...args);
  const text=(c.step as {text:string}).text,matches=selectedBindings.flatMap(n=>{const m=n.pattern.exec(text);return m?[{n,args:m.slice(1)}]:[];});assert.equal(matches.length,1,'Uniform step handler: '+text);return matches[0]!.n.run(c,...matches[0]!.args);
 }})),...extras];
}
