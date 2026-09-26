/** @script Deterministic native XML/ZIP/Office property campaign.
 * @usage bun scripts/property-campaign.ts [seed] [cases] [office-cases]
 * @description Writes exact corpus and reports under artifacts/property-campaign.
 */
import {mkdir} from 'node:fs/promises';
import {join,resolve} from 'node:path';
import {isDeepStrictEqual} from 'node:util';
import {Document,Presentation,Workbook,OoxmlError} from '../src/index.ts';
import {parseXml,applyEdits,escapeText,escapeAttribute,attribute,elements} from '../src/xml/index.ts';
import {readZip,writeZip} from '../src/opc/zip.ts';

export interface Input {text:string;attribute:string;replacement:string;bytes:number[];numbers:number[];}
interface Options {seed:number;cases:number;officeCases:number;checkpoint?:(caseId:string)=>void;}
const check=(ok:unknown,message:string)=>{if(!ok)throw Error(message);};
function budgets(seed:number,cases:number){if(!Number.isInteger(seed)||seed<0||seed>0xffffffff||!Number.isInteger(cases)||cases<1||cases>2000)throw Error('Invalid seed/case budget');}
export function makeInputs(seed:number,cases:number):Input[]{
 budgets(seed,cases);let state=seed>>>0;
 const random=()=>{state=(Math.imul(state,1664525)+1013904223)>>>0;return state;};
 const chars=['a','Z','0',' ','\t','\n','\r','&','<','>','"',"'",'π','中','é','😀'];
 const text=()=>{const size=random()%60;return Array.from({length:size},()=>chars[(random()>>>8)%chars.length]!).join('');};
 return Array.from({length:cases},()=>({text:text(),attribute:text(),replacement:text(),bytes:Array.from({length:random()%256},()=>random()>>>24),numbers:Array.from({length:8},()=>random())}));
}
const hash=(bytes:Uint8Array|string)=>new Bun.CryptoHasher('sha256').update(bytes).digest('hex');
const errorText=(e:unknown)=>e instanceof Error?e.stack??e.message:String(e);
function packageSnapshot(pkg:{names():string[];get(name:string):Uint8Array|undefined}){return new Map(pkg.names().map(n=>[n,pkg.get(n)!]));}
function custody(before:Map<string,Uint8Array>,after:{names():string[];get(name:string):Uint8Array|undefined},allowed:string[]){
 check(isDeepStrictEqual([...before.keys()].sort(),after.names().sort()),'Member names changed');
 for(const [n,bytes]of before)if(!allowed.includes(n))check(isDeepStrictEqual(after.get(n),bytes),'Unrelated member changed: '+n);
}
function refusal(fn:()=>unknown,code:string){let error:unknown;try{fn();}catch(e){error=e;}check(error instanceof OoxmlError&&error.code===code,'Expected '+code+', received '+String(error));}
const limits={maxEntries:64,maxEntryBytes:65536,maxTotalBytes:131072,maxArchiveBytes:262144,maxCompressionRatio:1000};

export function inspectZipMutation(mutant:Uint8Array,reader:typeof readZip=readZip):{opened?:Map<string,Uint8Array>;refusalCode?:string}{
 const before=mutant.slice();
 try{return {opened:reader(mutant,limits)};}catch(e){if(!(e instanceof OoxmlError)||!e.code.startsWith('zip-'))throw e;return {refusalCode:e.code};}
 finally{check(isDeepStrictEqual(mutant,before),'ZIP mutant caller buffer changed');}
}

export async function runCampaign(options:Options){
 budgets(options.seed,options.cases);if(!Number.isInteger(options.officeCases)||options.officeCases<1||options.officeCases>Math.min(100,options.cases))throw Error('Invalid Office case budget');
 const inputs=makeInputs(options.seed,options.cases),corpus=JSON.stringify(inputs),counts={xmlValid:0,xmlRefused:0,zipValid:0,zipMutations:0,zipAccepted:0,zipRefused:0,docx:0,pptx:0,xlsx:0,staleRefused:0};
 const failures:Array<{caseId:string;input:unknown;error:string}>=[],zipRefusalCodes:Record<string,number>={},durationsMs:Record<string,number>={};
 async function run(caseId:string,input:unknown,fn:()=>void|Promise<void>){try{await fn();options.checkpoint?.(caseId);}catch(error){failures.push({caseId,input,error:errorText(error)});}}
 let start=performance.now();
 for(const [i,input]of inputs.entries())await run(`xml/${i}`,input,()=>{
  const prefix=['p','q','Ω'][i%3]!,name=['leaf','Δ','中'][i%3]!;
  const source=`<?xml version="1.0"?><!--before--><r xmlns:${prefix}="urn:seed" a="${escapeAttribute(input.attribute)}"><${prefix}:${name}>${escapeText(input.text)}</${prefix}:${name}><keep attr="untouched"/></r><!--after-->`;
  const doc=parseXml(source),leaf=doc.root.children[0]!;
  check(doc.root.attributes.a===input.attribute&&leaf.text===input.text,'XML decoded text/attribute mismatch');
  check(leaf.namespaceURI==='urn:seed'&&leaf.localName===name,'Expanded name mismatch');
  check(attribute(doc.root,'a')===input.attribute,'Attribute lookup mismatch');
  const next=applyEdits(source,[{start:leaf.openEnd,end:leaf.closeStart,value:escapeText(input.replacement)}]);
  check(next===source.slice(0,leaf.openEnd)+escapeText(input.replacement)+source.slice(leaf.closeStart),'XML lexical custody mismatch');
  check(parseXml(next).root.children[0]!.text===input.replacement,'Accepted edit failed reparse');
  check(applyEdits(source,[])===source,'No-op changed source');counts.xmlValid++;
  refusal(()=>applyEdits(source,[{start:-1,end:0,value:''}]),'XML_EDIT_BOUNDS');counts.xmlRefused++;
  refusal(()=>applyEdits(source,[{start:0,end:3,value:''},{start:1,end:2,value:''}]),'XML_EDIT_OVERLAP');counts.xmlRefused++;
  refusal(()=>applyEdits(source,[{start:leaf.openEnd,end:leaf.closeStart,value:'<!DOCTYPE forbidden>'}]),'XML_EDIT_UNSAFE');counts.xmlRefused++;
 });durationsMs.xml=performance.now()-start;start=performance.now();
 for(const [i,input]of inputs.entries())for(const forceZip64 of [false,true])await run(`zip/${i}/${forceZip64?'64':'32'}`,input,()=>{
  const members=new Map([['folder/π.txt',new TextEncoder().encode(input.text)],['payload.bin',Uint8Array.from(input.bytes)],['empty.bin',new Uint8Array()]]);
  const archive=writeZip(members,{forceZip64}),original=archive.slice(),decoded=readZip(archive,limits);
  check(isDeepStrictEqual([...decoded],[...members].sort(([a],[b])=>a.localeCompare(b))), 'ZIP member payload roundtrip mismatch');
  check(isDeepStrictEqual(writeZip(members,{forceZip64}),archive),'ZIP output not deterministic');counts.zipValid++;
  for(let m=0;m<4;m++){
   const mutant=archive.slice(),position=input.numbers[m]!%mutant.length,bit=1<<(input.numbers[m+4]!%8);mutant[position]=mutant[position]!^bit;
   const id=`zip/${i}/${forceZip64?'64':'32'}/mutation/${m}`,detail={...input,position,bit,archiveSha256:hash(archive),mutantSha256:hash(mutant)};
   // A corruption may touch tolerated metadata. Admission itself is not a failure;
   // accepted data must roundtrip. Unexpected errors are never counted as refusals.
   try{
    counts.zipMutations++;const result=inspectZipMutation(mutant);
    if(result.refusalCode){counts.zipRefused++;zipRefusalCodes[result.refusalCode]=(zipRefusalCodes[result.refusalCode]??0)+1;continue;}
    const opened=result.opened!;
    check(isDeepStrictEqual([...readZip(writeZip(opened),limits)],[...opened].sort(([a],[b])=>a.localeCompare(b))),'Accepted mutant did not roundtrip');counts.zipAccepted++;
   }catch(error){failures.push({caseId:id,input:detail,error:errorText(error)});}
  }
  check(isDeepStrictEqual(archive,original),'ZIP input mutated');
 });durationsMs.zip=performance.now()-start;start=performance.now();
 for(const [i,input]of inputs.slice(0,options.officeCases).entries()){
  await run(`docx/${i}`,input,async()=>{
   const d=Document.create();d.addParagraph(`Before TOKEN after ${input.text}`);d.addParagraph('Untouched paragraph');
   const before=new Map(d.package.parts),p=d.paragraphs[0]!,span=p.find('TOKEN')[0]!;await span.replace(input.replacement);
   const target=`Before ${input.replacement} after ${input.text}`;check(d.paragraphs[0]!.text===target,'DOCX target mismatch');
   d.paragraphs[0]!.setRunFormatting({bold:Boolean(i%2),italic:Boolean(i%3)});const saved=await d.save(),q=await Document.open(saved);
   check(q.paragraphs[0]!.text===target&&q.paragraphs[1]!.text==='Untouched paragraph','DOCX saved text mismatch');
   check(isDeepStrictEqual([...q.package.parts.keys()].sort(),[...before.keys()].sort()),'DOCX member names changed');
   for(const [n,b]of before)if(n!=='word/document.xml')check(isDeepStrictEqual(q.package.get(n),b),'DOCX unrelated member changed');
   let error:unknown;try{await span.replace('stale');}catch(e){error=e;}check(error instanceof OoxmlError&&error.code==='docx-stale-span','DOCX stale span was accepted');counts.staleRefused++;check(isDeepStrictEqual(await d.save(),saved),'DOCX stale refusal changed archive');counts.docx++;
  });
  await run(`pptx/${i}`,input,async()=>{
   const d=Presentation.create();d.addTextSlide('Before TOKEN');d.addTextSlide('Untouched slide');const s=d.slides[0]!,before=packageSnapshot(d.package),old=s.inspectText('old')[0]!.anchor;
   const text=input.text.replace(/\r\n?/g,'\n'),receipt=s.addTextBox(input.text,{x:i*100,y:i*200,width:914400,height:914400},{bold:Boolean(i%2),italic:false});
   const saved=d.package.toBytes(),q=await Presentation.open(saved);check(isDeepStrictEqual(q.slides[0]!.inspectText('saved').slice(1).map(p=>p.text),text.split('\n')),'PPTX text box mismatch');
   check(receipt.paragraphCount===text.split('\n').length,'PPTX paragraph receipt mismatch');custody(before,q.package,[s.partName]);
   refusal(()=>s.replaceTextAt(old,'TOKEN','stale'),'PPTX_STALE_ANCHOR');counts.staleRefused++;check(isDeepStrictEqual(d.package.toBytes(),saved),'PPTX stale refusal changed archive');counts.pptx++;
  });
  await run(`xlsx/${i}`,input,async()=>{
   const w=Workbook.create();w.worksheet('Sheet1').setCellValue('A1',input.text);w.addWorksheet('Other').setCellValue('B2',input.numbers[0]!%1000);const before=packageSnapshot(w.package),sheet=w.worksheet('Sheet1');sheet.setCellStyle('A1',0);
   const q=await Workbook.open(w.toBytes());check(q.worksheet('Sheet1').getCell('A1')?.value===input.text&&q.worksheet('Sheet1').getCell('A1')?.styleId==='0','XLSX saved value/style mismatch');custody(before,q.package,['xl/worksheets/sheet1.xml']);
   const saved=w.toBytes();refusal(()=>sheet.setCellStyle('A1',999),'xlsx-style-invalid');check(isDeepStrictEqual(w.toBytes(),saved),'XLSX invalid style changed archive');sheet.setCellStyle('A1',null);const removed=await Workbook.open(w.toBytes());check(removed.worksheet('Sheet1').getCell('A1')?.styleId===undefined&&removed.worksheet('Sheet1').getCell('A1')?.value===input.text,'XLSX removal saved state mismatch');check(removed.worksheet('Other').getCell('B2')?.value===input.numbers[0]!%1000,'XLSX other sheet changed');custody(before,removed.package,['xl/worksheets/sheet1.xml']);counts.xlsx++;
  });
 }
 durationsMs.office=performance.now()-start;
 return {schemaVersion:1,status:failures.length?'failed':'passed',seed:options.seed,cases:options.cases,officeCases:options.officeCases,inputSha256:hash(corpus),counts,zipRefusalCodes,durationsMs,failures,limits:{xmlTextSymbols:59,payloadBytes:255,zipReader:limits,mutation:'four deterministic single-bit flips per ZIP32/ZIP64 archive',source:'synthetic; no shared fixtures modified',coverage:'sampled deterministic properties, not exhaustive fuzzing or schema validation'}};
}

if(import.meta.main){
 const options={seed:Number(process.argv[2]??'20260926'),cases:Number(process.argv[3]??'128'),officeCases:Number(process.argv[4]??'16')};
 const root=resolve(import.meta.dir,'..'),out=join(root,'artifacts/property-campaign',`seed-${options.seed}`);await mkdir(out,{recursive:true});
 const runId=crypto.randomUUID();await Bun.write(join(out,'report.json'),JSON.stringify({runId,status:'running'}));
 try{
  const report=await runCampaign(options);await Bun.write(join(out,'corpus.json'),JSON.stringify(makeInputs(options.seed,options.cases)));
  const runtimeSources:Record<string,string>={};for await(const file of new Bun.Glob('src/**/*.ts').scan({cwd:root}))runtimeSources[file]=hash(await Bun.file(join(root,file)).bytes());
  await Bun.write(join(out,'report.json'),JSON.stringify({...report,runId,sourceSha256:hash(await Bun.file(import.meta.path).bytes()),runtimeSources,versions:{bun:Bun.version,platform:process.platform,arch:process.arch}},null,2)+'\n');
  console.log(JSON.stringify({status:report.status,seed:report.seed,counts:report.counts,failures:report.failures},null,2));if(report.status!=='passed')process.exitCode=1;
 }catch(error){await Bun.write(join(out,'report.json'),JSON.stringify({runId,status:'failed',error:errorText(error),options},null,2));throw error;}
}
