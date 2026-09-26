import { lstat, realpath } from "node:fs/promises";
import { basename, dirname, extname, join, resolve } from "node:path";
import { Document } from "../docx/index.ts";
import { trackedReplace } from '../docx/redline.ts';
import { inspectStories } from '../docx/story.ts';
import { inspectRevisions } from '../docx/revisions.ts';
import { Presentation } from "../pptx/index.ts";
import { findPlaceholderText } from "../pptx/placeholders.ts";
import { Workbook } from "../xlsx/index.ts";
import { setCellWrapText } from "../xlsx/styles.ts";
import { OpcPackage } from "../opc/package.ts";
import { OoxmlError } from "../errors.ts";
import { parseXml,elements } from "../xml/index.ts";

export type PatchValue=string|number|boolean|null;
export interface PatchRequest {
  source:string; output?:string; format?:"docx"|"pptx"|"xlsx";
  mode:"dry_run"|"strict"|"safe";
  changes:{target:string;value:PatchValue}[];
  multilineWrap?:boolean;
  trackChanges?:boolean;
  revisionMetadata?:{author:string;date:string};
  calculationPolicy?:"invalidate-without-recalculation";
  expectedSourceSha256?:string;
  expectedDestinationSha256?:string|null;
}
export interface TargetResult {
  target:string;value:PatchValue;matched:number;
  status:"matched"|"unmatched"|"unsupported"|"unchanged"|"committed";
  code?:string;
}
export interface PatchReceipt {
  status:"preview"|"refused"|"committed"; committedChanges:number;
  results:TargetResult[]; sourceSha256:string;outputSha256?:string;
  calculationState:"unchanged"|"recalculation-required";
  trackedRevisions:number; previewRevisions:number; revisionIds:string[];
  error?:{code:string;message:string};changedParts:string[];
}
interface Snapshot {path:string; bytes?:Uint8Array; hash:string|null; dev?:number;ino?:number;}
interface Target {key:string;result:TargetResult;invalidatesCalculation?:boolean;range?:{paragraph:number;start:number;end:number};apply:()=>boolean|Promise<boolean>; verify:(bytes:Uint8Array)=>Promise<void>;}
const locks=new Map<string,Promise<void>>();
const sha=(b:Uint8Array)=>new Bun.CryptoHasher("sha256").update(b).digest("hex");
function refuse(code:string,message:string):never {throw new OoxmlError(code,message);}

/** Resolve and stage a complete edit batch in memory. Dry runs never write.
 * Only this API's writers are serialised; caller directories must be trusted.
 * External editors can still race the final fingerprint check and atomic rename.
 * This first slice is all-targets-required in every mode, with no best-effort fallback.
 */
export async function patchOffice(request:PatchRequest):Promise<PatchReceipt> {
  const receipt:PatchReceipt={status:"refused",committedChanges:0,results:[],sourceSha256:"",calculationState:"unchanged",changedParts:[],trackedRevisions:0,previewRevisions:0,revisionIds:[]};
  try {
    validateRequest(request);
    // Own the request before the first await; caller mutation must not change
    // tracked intent, author/date or any resolved target during asynchronous I/O.
    request={...request,changes:request.changes.map(c=>({target:c.target,value:c.value})),revisionMetadata:request.revisionMetadata?{author:request.revisionMetadata.author,date:request.revisionMetadata.date}:undefined};
    // Resolve parent links so aliases share a writer lock; refuse leaf links explicitly.
    const sourcePath=await canonicalPath(request.source);
    const outputPath=request.output?await canonicalPath(request.output):sourcePath;
    const paths=[...new Set([sourcePath,outputPath])].sort();
    return await withLocks(paths,async()=>{
      try {
        const source=await snapshot(sourcePath);
        if(!source.bytes)refuse("workflow-source-missing","Source file does not exist");
        receipt.sourceSha256=source.hash!;
        const destination=outputPath===sourcePath?source:await snapshot(outputPath);
        if(request.expectedSourceSha256!==undefined&&request.expectedSourceSha256!==source.hash)refuse("workflow-stale-source","Source fingerprint differs from requested revision");
        if(request.expectedDestinationSha256!==undefined&&request.expectedDestinationSha256!==destination.hash)refuse("workflow-stale-destination","Destination fingerprint differs from requested revision");
        const sameFile=sourcePath===outputPath||!!destination.bytes&&(source.dev===destination.dev&&source.ino===destination.ino);
        if(request.mode==="safe"&&(!request.output||sameFile))refuse("workflow-output-alias","Safe mode requires a distinct destination");
        if(sourcePath!==outputPath&&sameFile)refuse("workflow-output-alias","Hardlink output aliases refuse");
        const inferred=extname(sourcePath).slice(1).toLowerCase();
        const format=request.format??inferred;
        if(!["docx","pptx","xlsx"].includes(format))refuse("workflow-format-unsupported","Only DOCX/PPTX/XLSX are supported");
        if(request.multilineWrap&&format!=="xlsx")refuse("workflow-option-unsupported","Wrapping applies to XLSX only");
        if(request.calculationPolicy&&format!=="xlsx")refuse("workflow-option-unsupported","Calculation policy applies to XLSX only");
        if(request.trackChanges&&format!=="docx")refuse('workflow-option-unsupported','Tracked replacement applies to DOCX only');
        const original=await OpcPackage.open(source.bytes);
        const tracked=request.trackChanges?{pkg:await OpcPackage.open(source.bytes),ids:[] as string[]}:undefined;
        const doc=format==="docx"&&!tracked?await Document.open(source.bytes):undefined;
        const deck=format==="pptx"?await Presentation.open(source.bytes):undefined;
        const book=format==="xlsx"?await Workbook.open(source.bytes):undefined;
        const targets:Target[]=[];
        const keys=new Set<string>();
        const expectedDoc=doc?.paragraphs.map(p=>p.text);
        // Every target is resolved before any mutation. A later failure cannot
        // turn an earlier match into a committed or partially applied operation.
        for(const change of request.changes) {
          const result:TargetResult={...change,matched:0,status:"matched"};receipt.results.push(result);
          try {
            const target=tracked?resolveTrackedDocx(tracked,change,result,request.revisionMetadata!):doc?resolveDocx(doc,change,result):deck?resolvePptx(deck,change,result):resolveXlsx(book!,change,result,request.multilineWrap===true);
            if(keys.has(target.key))refuse("workflow-overlapping-targets",`Repeated target ${change.target}`);
            if(target.range&&targets.some(t=>t.range&&t.range.paragraph===target.range!.paragraph&&t.range.start<target.range!.end&&target.range!.start<t.range.end))refuse("workflow-overlapping-targets","Overlapping text ranges refuse");
            keys.add(target.key);targets.push(target);
          } catch(error) {
            result.code=errorCode(error);result.status=result.code==="workflow-target-missing"?"unmatched":"unsupported";
            if(!receipt.error)receipt.error={code:result.code,message:errorMessage(error)};
          }
        }
        if(receipt.error)return receipt;
        if(expectedDoc)for(const t of [...targets].sort((a,b)=>(b.range?.start??0)-(a.range?.start??0))) {
          const r=t.range!;const text=expectedDoc[r.paragraph]!;
          expectedDoc[r.paragraph]=text.slice(0,r.start)+t.result.value+text.slice(r.end);
        }
        const changed:boolean[]=[];
        for(const target of targets)changed.push(await target.apply());
        const outputBytes=tracked?tracked.pkg.toBytes():doc?await doc.save():deck?deck.package.toBytes():book!.package.toBytes();
        // Reopen once for package invariants and verify each requested effect.
        const outputPackage=await OpcPackage.open(outputBytes);
        for(const target of targets)await target.verify(outputBytes);
        if(expectedDoc&&JSON.stringify((await Document.open(outputBytes)).paragraphs.map(p=>p.text))!==JSON.stringify(expectedDoc))refuse("workflow-verification-failed","DOCX saved paragraph outcomes differ from the resolved batch");
        receipt.changedParts=partChanges(original,outputPackage);
        const calculationState=book&&targets.some(t=>t.invalidatesCalculation)&&hasFormulas(book)?"recalculation-required":"unchanged";
        if(request.mode==="dry_run") {
          receipt.status="preview";
          receipt.previewRevisions=tracked?.ids.length??0;
          // Preview describes the change without claiming the source now needs recalculation.
          return receipt;
        }
        await checkSnapshot(source,"workflow-stale-source");
        if(destination.path!==source.path)await checkSnapshot(destination,"workflow-stale-destination");
        await outputPackage.save(outputPath);
        receipt.status="committed";
        receipt.outputSha256=sha(outputBytes);
        receipt.calculationState=calculationState;
        receipt.results.forEach((result,i)=>{result.status=changed[i]?"committed":"unchanged";});
        receipt.committedChanges=changed.filter(Boolean).length;
        receipt.trackedRevisions=tracked?.ids.length??0;
        receipt.revisionIds=tracked?.ids.slice()??[];
        return receipt;
      } catch(error) {return failedReceipt(receipt,error);}
    });
  } catch(error) {return failedReceipt(receipt,error);}
}

function resolveTrackedDocx(state:{pkg:OpcPackage;ids:string[]},change:PatchRequest['changes'][number],result:TargetResult,metadata:{author:string;date:string}):Target {
  if(typeof change.value!=='string')refuse('workflow-value-unsupported','DOCX replacements require string values');
  const part=state.pkg.mainPart(),replacement=change.value;
  const story=inspectStories(state.pkg,{view:'current'}).stories.find(s=>s.part===part);
  let matches=0;
  for(const p of story?.paragraphs??[])for(let index=0;(index=p.text.indexOf(change.target,index))>=0;index++)matches++;
  result.matched=matches;
  if(matches!==1)refuse(matches?'workflow-target-ambiguous':'workflow-target-missing','Tracked replacement requires one exact main-story target');
  return {key:change.target,result,apply:()=>{
    const applied=trackedReplace(state.pkg,part,change.target,replacement,metadata);state.ids=applied.revisionIds;return state.ids.length>0;
  },verify:async bytes=>{
    const reopened=await OpcPackage.open(bytes),revisions=inspectRevisions(reopened);
    const observed=revisions.revisions.filter(r=>r.part===part);
    if(revisions.unsupported.some(r=>r.part===part)||observed.length!==state.ids.length||observed.some(r=>!state.ids.includes(r.id)||r.author!==metadata.author||r.date!==metadata.date))refuse('workflow-verification-failed','Saved tracked revisions differ from staged revision identities');
  }};
}
function resolveDocx(doc:Document,change:PatchRequest["changes"][number],result:TargetResult):Target {
  if(typeof change.value!=="string")refuse("workflow-value-unsupported","DOCX replacements require string values");
  const replacement=change.value;
  const matches=doc.find(change.target);result.matched=matches.length;
  if(!matches.length)refuse("workflow-target-missing",`No match for ${change.target}`);
  if(matches.length!==1)refuse("workflow-target-ambiguous",`More than one match for ${change.target}`);
  const paragraph=doc.paragraphs.findIndex(p=>p.text.includes(change.target));
  if(paragraph<0)refuse("workflow-target-missing","No supported paragraph");
  const start=doc.paragraphs[paragraph]!.text.indexOf(change.target);
  return {key:change.target,result,range:{paragraph,start,end:start+change.target.length},apply:async()=>{
    if(change.target===replacement)return false;
    const current=doc.find(change.target);
    if(current.length!==1)refuse("workflow-target-changed","Earlier edit changed a later target");
    await current[0]!.replace(replacement);return true;
  },verify:async bytes=>{
    const after=await Document.open(bytes);
    if(replacement&&!after.paragraphs.some(p=>p.text.includes(replacement)))refuse("workflow-verification-failed","DOCX requested replacement missing after save");
  }};
}
function resolvePptx(deck:Presentation,change:PatchRequest["changes"][number],result:TargetResult):Target {
  const match=/^slide:([1-9]\d*)\/(title|subtitle)$/.exec(change.target);
  if(!match)refuse("workflow-target-unsupported","Use slide:N/title or slide:N/subtitle");
  if(typeof change.value!=="string")refuse("workflow-value-unsupported","PPTX replacements require string values");
  const value=change.value,index=Number(match[1])-1,kind=match[2] as "title"|"subtitle";
  const slide=deck.slides[index];if(!slide)refuse("workflow-target-missing","Slide does not exist");
  let resolved:ReturnType<typeof findPlaceholderText>;
  try{resolved=findPlaceholderText(slide,kind);}catch(error){if(errorCode(error)==="PPTX_PLACEHOLDER_NOT_FOUND")refuse("workflow-target-missing","Placeholder missing");throw error;}
  result.matched=1;
  return {key:`slide:${index}/${kind}`,result,apply:()=>{
    if(value===resolved.text)return false;
    const current=findPlaceholderText(slide,kind);
    if(current.text!==resolved.text)refuse("workflow-target-changed","Earlier edit changed the placeholder");
    slide.replaceTextAt(current.anchor,current.text,value);return true;
  },verify:async bytes=>{
    const after=await Presentation.open(bytes);
    // Empty result is valid, even though it is not an editable future placeholder.
    const texts=after.slides[index]!.inspectText("");
    if(texts[resolved.anchor.paragraphIndex]?.text!==value)refuse("workflow-verification-failed","PPTX requested text missing after save");
  }};
}
function resolveXlsx(book:Workbook,change:PatchRequest["changes"][number],result:TargetResult,wrap:boolean):Target {
  const split=change.target.lastIndexOf("!");
  const sheetName=split>=0?change.target.slice(0,split):book.sheetnames[0]!;
  const address=(split>=0?change.target.slice(split+1):change.target).toUpperCase();
  if(!/^[A-Z]{1,3}[1-9]\d*$/.test(address))refuse("workflow-target-unsupported","Use Sheet!A1 or A1 cell targets");
  if(!book.sheetnames.includes(sheetName))refuse("workflow-target-missing",`Worksheet ${sheetName} does not exist`);
  const sheet=book.worksheet(sheetName),cell=sheet.getCell(address);
  if(!cell)refuse("workflow-target-missing",`Cell ${address} does not exist`);
  result.matched=1;
  if(cell.kind==="formula")refuse("workflow-target-unsupported","Formula overwrite requires the formula API");
  if(change.value===null)refuse("workflow-value-unsupported","Clearing cells is not yet supported by this workflow");
  const value=change.value;
  return {key:`${sheetName}!${address}`,result,invalidatesCalculation:!Object.is(cell.value,value),apply:()=>{
    const before=book.package.toBytes();
    if(!Object.is(cell.value,value))book.worksheet(sheetName).setCellValue(address,value);
    if(wrap&&typeof change.value==="string"&&change.value.includes("\n"))setCellWrapText(book,sheetName,address,true);
    return sha(book.package.toBytes())!==sha(before);
  },verify:async bytes=>{
    const after=await Workbook.open(bytes);
    if(!Object.is(after.worksheet(sheetName).getCell(address)?.value,change.value))refuse("workflow-verification-failed","XLSX requested value missing after save");
  }};
}
function hasFormulas(book:Workbook):boolean{return book.package.relationships(book.package.mainPart()).filter(r=>!r.external&&r.type.endsWith('/worksheet')).some(r=>elements(parseXml(book.package.text(r.resolved!)), 'f', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main').length>0);}
function partChanges(before:OpcPackage,after:OpcPackage):string[]{return [...new Set([...before.names(),...after.names()])].filter(name=>{const a=before.get(name),b=after.get(name);return !a||!b||sha(a)!==sha(b);}).sort();}
function validateRequest(r:PatchRequest):void {
  if(!r||typeof r.source!=="string"||!r.source||!["dry_run","strict","safe"].includes(r.mode)||!Array.isArray(r.changes)||!r.changes.length)refuse("workflow-request-invalid","Source, mode and nonempty changes required");
  if(r.trackChanges!==undefined&&typeof r.trackChanges!=='boolean')refuse('workflow-request-invalid','trackChanges must be boolean');
  if(r.trackChanges){
    if(r.changes.length!==1)refuse('workflow-tracked-batch-unsupported','Tracked workflow supports exactly one target');
    if(!r.revisionMetadata||typeof r.revisionMetadata.author!=='string'||typeof r.revisionMetadata.date!=='string')refuse('workflow-request-invalid','Tracked changes require explicit author and UTC date');
  }
  if(r.changes.length>1000)refuse("workflow-request-limit","At most 1000 targets per batch");
  if(r.calculationPolicy!==undefined&&r.calculationPolicy!=="invalidate-without-recalculation")refuse("workflow-policy-unsupported","Only cache invalidation without recalculation is available");
  for(const c of r.changes)if(!c||typeof c.target!=="string"||!c.target||!(c.value===null||typeof c.value==="string"||typeof c.value==="boolean"||typeof c.value==="number"&&Number.isFinite(c.value)))refuse("workflow-request-invalid","Invalid target/value");
}
async function canonicalPath(path:string):Promise<string>{const absolute=resolve(path);try{if((await lstat(absolute)).isSymbolicLink())refuse("workflow-symlink","Leaf symlink paths refuse");}catch(e){if((e as NodeJS.ErrnoException).code!=="ENOENT")throw e;}return join(await realpath(dirname(absolute)),basename(absolute));}
async function snapshot(path:string):Promise<Snapshot>{
  try {const stat=await lstat(path);if(stat.isSymbolicLink())refuse("workflow-symlink","Leaf symlink paths refuse");if(!stat.isFile())refuse("workflow-path-invalid","Input/output must be regular files");if(stat.size>256*1024*1024)refuse("workflow-file-limit","File exceeds 256 MiB");const bytes=await Bun.file(path).bytes();return {path,bytes,hash:sha(bytes),dev:stat.dev,ino:stat.ino};}
  catch(e){if((e as NodeJS.ErrnoException).code==="ENOENT")return {path,hash:null};throw e;}
}
async function checkSnapshot(before:Snapshot,code:string):Promise<void>{const now=await snapshot(before.path);if(now.hash!==before.hash||now.dev!==before.dev||now.ino!==before.ino)refuse(code,"File changed while staging the batch");}
function failedReceipt(receipt:PatchReceipt,error:unknown):PatchReceipt{receipt.status="refused";receipt.committedChanges=0;receipt.changedParts=[];receipt.calculationState="unchanged";receipt.trackedRevisions=0;receipt.previewRevisions=0;receipt.revisionIds=[];receipt.error={code:errorCode(error),message:errorMessage(error)};return receipt;}
function errorCode(error:unknown):string{return error instanceof OoxmlError?error.code:"workflow-io-or-validation-failed";}
function errorMessage(error:unknown):string{return error instanceof Error?error.message:String(error);}
async function withLocks<T>(paths:string[],fn:()=>Promise<T>):Promise<T>{
 const releases:(()=>void)[]=[];
 try{for(const path of paths){const prior=locks.get(path)??Promise.resolve();let release!:()=>void;const current=new Promise<void>(r=>{release=r;});const tail=prior.then(()=>current);locks.set(path,tail);await prior;releases.push(()=>{release();if(locks.get(path)===tail)locks.delete(path);});}return await fn();}
 finally{for(const release of releases.reverse())release();}
}
