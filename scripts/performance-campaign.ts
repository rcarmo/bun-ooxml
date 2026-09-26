/** @script Bounded native workloads with raw timing samples; no speedup claim.
 * @usage bun scripts/performance-campaign.ts
 * @description Generates fixed in-memory inputs and reports median/p95, no timing pass thresholds.
 */
import {mkdir} from 'node:fs/promises';
import {join,resolve} from 'node:path';
import {cpus,platform,arch} from 'node:os';
import {Document,Presentation,Workbook} from '../src/index.ts';
import {parseXml,applyEdits,elements} from '../src/xml/index.ts';
import {readZip,writeZip} from '../src/opc/zip.ts';

export function timingStats(samples:number[]){
 if(!samples.length||samples.some(n=>!Number.isFinite(n)||n<0))throw Error('Invalid timing samples');
 const sorted=[...samples].sort((a,b)=>a-b),middle=Math.floor(sorted.length/2);
 return {samplesMs:[...samples],medianMs:sorted.length%2?sorted[middle]!:(sorted[middle-1]!+sorted[middle]!)/2,p95Ms:sorted[Math.ceil(sorted.length*.95)-1]!,minMs:sorted[0]!,maxMs:sorted.at(-1)!};
}
const check=(value:unknown,message:string)=>{if(!value)throw Error(message);};
const hash=(b:Uint8Array|string)=>new Bun.CryptoHasher('sha256').update(b).digest('hex');

export async function runPerformance(){
 const workloads:Array<{id:string;input:Uint8Array|string;dimensions:Record<string,number>;run:()=>Promise<void>|void}>=[];
 const xml='<r>'+Array.from({length:1000},(_,i)=>`<n id="${i}">text ${i}</n>`).join('')+'</r>';
 workloads.push({id:'xml-parse-edit-reparse',input:xml,dimensions:{elements:1001},run:()=>{const d=parseXml(xml),node=d.root.children[500]!,out=applyEdits(xml,[{start:node.openEnd,end:node.closeStart,value:'changed'}]);check(parseXml(out).root.children[500]!.text==='changed','XML output mismatch');}});
 const zipParts=new Map(Array.from({length:128},(_,i)=>[`part${i.toString().padStart(3,'0')}.bin`,Uint8Array.from({length:4096},(_,j)=>(j*31+i)%256)] as const));
 const zip=writeZip(zipParts,{forceZip64:true});
 workloads.push({id:'zip64-read-write-read',input:zip,dimensions:{members:128,payloadBytes:128*4096},run:()=>{const parts=readZip(zip),reopened=readZip(writeZip(parts,{forceZip64:true}));check(reopened.size===128,'ZIP member count mismatch');for(const [n,b]of zipParts)check(hash(reopened.get(n)!)===hash(b),'ZIP payload mismatch');}});
 const doc=Document.create();for(let i=0;i<50;i++)doc.addParagraph(`Paragraph ${i} token${i}`);const docBytes=await doc.save();
 workloads.push({id:'docx-open-edit-save-reopen',input:docBytes,dimensions:{paragraphs:50,edits:1},run:async()=>{const d=await Document.open(docBytes);await d.paragraphs[25]!.find('token25')[0]!.replace('changed');const q=await Document.open(await d.save());check(q.paragraphs.length===50&&q.paragraphs[25]!.text==='Paragraph 25 changed','DOCX output mismatch');}});
 const deck=Presentation.create();for(let i=0;i<8;i++)deck.addTextSlide(`Slide ${i}`);const pptBytes=deck.package.toBytes();
 workloads.push({id:'pptx-open-append-box-save-reopen',input:pptBytes,dimensions:{slides:8,boxesAdded:1},run:async()=>{const d=await Presentation.open(pptBytes);d.slides[3]!.addTextBox('Measured box',{x:0,y:0,width:914400,height:914400});const q=await Presentation.open(d.package.toBytes());check(q.slides.length===8&&q.slides[3]!.inspectText('check').some(p=>p.text==='Measured box'),'PPTX output mismatch');}});
 const book=Workbook.create(),sheet=book.worksheet('Sheet1');for(let i=1;i<=128;i++)sheet.setCellValue(`A${i}`,i);const xlsBytes=book.toBytes();
 workloads.push({id:'xlsx-open-style-save-reopen',input:xlsBytes,dimensions:{cells:128,styleEdits:1},run:async()=>{const w=await Workbook.open(xlsBytes);w.worksheet('Sheet1').setCellStyle('A64',0);const q=await Workbook.open(w.toBytes());check(q.worksheet('Sheet1').getCell('A64')?.styleId==='0'&&q.worksheet('Sheet1').getCell('A64')?.value===64,'XLSX output mismatch');}});
 const results=[];for(const w of workloads){await w.run();const samples:number[]=[];for(let i=0;i<7;i++){const start=performance.now();await w.run();samples.push(performance.now()-start);}results.push({id:w.id,inputBytes:typeof w.input==='string'?Buffer.byteLength(w.input):w.input.length,inputSha256:hash(w.input),dimensions:w.dimensions,warmups:1,...timingStats(samples),correctness:'passed'});}
 return {schemaVersion:1,status:'passed',environment:{bun:Bun.version,platform:platform(),arch:arch(),cpu:cpus()[0]?.model,logicalCpus:cpus().length},method:{timing:'in-memory open/mutate/serialize/reopen and listed correctness checks included; input construction excluded',samples:7,warmups:1,p95:'nearest rank',concurrency:1,gc:'not controlled',memory:'process RSS after all workloads; not peak or per-operation allocation',comparison:'none; no portable performance threshold'},rssAfterBytes:process.memoryUsage().rss,results};
}
export async function writePerformanceReport(path:string,run:()=>Promise<unknown>=runPerformance){
 const runId=crypto.randomUUID();await Bun.write(path,JSON.stringify({runId,status:'running'}));
 try{const report=await run();await Bun.write(path,JSON.stringify({...report as object,runId,sourceSha256:hash(await Bun.file(import.meta.path).bytes())},null,2)+'\n');return report;}
 catch(error){await Bun.write(path,JSON.stringify({runId,status:'failed',error:error instanceof Error?error.stack:String(error)},null,2)+'\n');throw error;}
}
if(import.meta.main){const out=join(resolve(import.meta.dir,'..'),'artifacts/property-campaign');await mkdir(out,{recursive:true});const report=await writePerformanceReport(join(out,'performance.json')) as Awaited<ReturnType<typeof runPerformance>>;console.log(JSON.stringify(report.results.map(r=>({id:r.id,medianMs:r.medianMs,p95Ms:r.p95Ms})),null,2));}
