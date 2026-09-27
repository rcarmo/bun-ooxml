/** @script Independent development-only schema, PDF text and calculation smoke.
 * @usage bun scripts/office-oracles.ts
 * @description Requires dotnet10, LibreOffice24.2.7 and Poppler; never imported by src.
 */
import {mkdir,mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {Document,Presentation,Workbook,OpcPackage} from '../src/index.ts';
import {parseXml,elements} from '../src/xml/index.ts';
import {runOracleCommand} from './oracle-process.ts';

const root=resolve(import.meta.dir,'..'),out=join(root,'artifacts/office-oracles');
const sandbox=await mkdtemp(join(tmpdir(),'bun-office-oracles-'));
const sha=(bytes:Uint8Array)=>new Bun.CryptoHasher('sha256').update(bytes).digest('hex');
const report:any={schemaVersion:1,status:'running',sources:{},runtime:'Bun-only; external programs used here solely as test oracles',versions:{},commands:[],inputs:[],schema:null,negativeControl:null,rendering:[],calculation:null,limits:['Three authored samples, not all fixtures or format operations','PDF page/text assertions do not establish visual fidelity or font equivalence','LibreOffice is not Microsoft Office','One arithmetic formula recalculated externally; no Bun calculation credit']};
await mkdir(out,{recursive:true});
async function command(args:string[],seconds=90){
 const result=await runOracleCommand(args,root,seconds*1000);report.commands.push(result);
 if(result.timedOut||result.outputLimited)throw Error('Oracle timeout/output limit: '+args[0]);return result;
}
const check=(condition:unknown,message:string)=>{if(!condition)throw Error(message);};
async function ok(args:string[],seconds=90){const r=await command(args,seconds);check(r.exit===0,`${args[0]} failed: ${r.stderr}`);return r;}
try{
 for(const source of ['scripts/office-oracles.ts','scripts/oracle-process.ts','tests/oracles/schema/Program.cs','tests/oracles/schema/SchemaCheck.csproj','tests/oracles/schema/packages.lock.json','src/docx/index.ts','src/docx/style-authoring.ts','src/docx/page-layout.ts','src/docx/effective-formatting.ts','src/docx/row-header.ts','src/docx/paragraph-text.ts','src/docx/append-run.ts','src/pptx/index.ts','src/pptx/text-box.ts','src/pptx/slide-order.ts','src/xlsx/index.ts','src/xlsx/cell-style.ts'])report.sources[source]=sha(await Bun.file(join(root,source)).bytes());
 report.versions.dotnet=(await ok(['dotnet','--version'])).stdout.trim();
 report.versions.libreoffice=(await ok(['libreoffice','--version'])).stdout.trim();
 const poppler=await ok(['pdftotext','-v']);report.versions.poppler=(poppler.stderr||poppler.stdout).trim();
 const project='tests/oracles/schema/SchemaCheck.csproj';
 await ok(['dotnet','restore',project,'--locked-mode']);await ok(['dotnet','build',project,'--no-restore','--configuration','Release']);
 const validator=join(root,'tests/oracles/schema/bin/Release/net10.0/SchemaCheck.dll');
 const doc=Document.create();doc.addParagraphStyle('Smoke',{name:'Oracle heading',bold:true});doc.addParagraph('Native Word oracle',{style:'Smoke'});doc.addParagraph('Obsolete text.').appendRun(' trailing text').setText('A second paragraph.');
 doc.addParagraphStyle('CascadeBase',{name:'Cascade Base',bold:true});doc.addParagraphStyle('CascadeToggle',{name:'Cascade Toggle',basedOn:'CascadeBase',bold:true,italic:true});
 doc.addParagraph('BOLDMARK',{style:'CascadeBase'});doc.addParagraph('TOGGLEMARK',{style:'CascadeToggle'});doc.addParagraph('PLAINMARK',{style:'CascadeBase'});doc.paragraphs.at(-1)!.setRunFormatting({bold:false});doc.addParagraph('BOTHMARK',{bold:true,italic:true});
 const formattingExpected=doc.paragraphs.filter(p=>p.text.endsWith('MARK')).map(p=>{const r=p.effectiveRunFormatting()[0]!;return {text:r.text,bold:r.bold.value,italic:r.italic.value};});
 check(JSON.stringify(formattingExpected)===JSON.stringify([{text:'BOLDMARK',bold:true,italic:false},{text:'TOGGLEMARK',bold:false,italic:true},{text:'PLAINMARK',bold:false,italic:false},{text:'BOTHMARK',bold:true,italic:true}]),'Native cascade probe mismatch');
 const table=doc.addTable(2,2);table.cell(0,0).text='Answer';table.cell(0,1).text='42';table.setRowHeader(0,true);table.setRowHeader(1,false);check(table.isRowHeader(0)&&!table.isRowHeader(1),'Direct header readback mismatch');doc.setPageLayout({...doc.getPageLayout(),width:15840,height:12240,orientation:'landscape'});
 const deck=Presentation.create();deck.addTextSlide('Native slide oracle','First page');const slide=deck.addTextSlide('Second slide oracle');slide.addTextBox('Positioned oracle box',{x:914400,y:914400,width:5486400,height:914400},{bold:true});slide.addTable(1,2,{x:914400,y:2743200,width:5486400,height:914400});slide.tables[0]!.cell(0,0).text='Answer';slide.tables[0]!.cell(0,1).text='42';deck.reorderSlides([1,0]);
 const book=Workbook.create();book.worksheet('Sheet1').setCellValue('A1',20);book.worksheet('Sheet1').setCellValue('A2',22);book.worksheet('Sheet1').setCellValue('A3',999);book.worksheet('Sheet1').setCellStyle('A1',0);
 // Explicit oracle fixture assembly via native OPC; there is no formula-authoring API claim.
 const sheet='xl/worksheets/sheet1.xml',xml=book.package.text(sheet),cell=elements(parseXml(xml),'c','http://schemas.openxmlformats.org/spreadsheetml/2006/main').find(c=>c.attributes.r==='A3')!;
 book.package.set(sheet,xml.slice(0,cell.start)+'<c r="A3"><f>SUM(A1:A2)</f><v>999</v></c>'+xml.slice(cell.end));
 const main=book.package.mainPart();book.package.set(main,book.package.text(main).replace('</workbook>','<calcPr fullCalcOnLoad="1" forceFullCalc="1"/></workbook>'));
 const editedBook=await Workbook.open(book.toBytes());editedBook.worksheet('Sheet1').setCellValue('A1',21);
 check(editedBook.worksheet('Sheet1').getCell('A3')?.cached===null,'Bun value edit did not invalidate the old formula cache');
 const paths=[join(out,'word.docx'),join(out,'deck.pptx'),join(out,'book.xlsx')];
 await doc.save(paths[0]);await deck.save(paths[1]!);await Bun.write(paths[2]!,editedBook.toBytes());
 for(const path of paths)report.inputs.push({file:path.split('/').at(-1),sha256:sha(await Bun.file(path).bytes())});
 const validation=await command(['dotnet',validator,...paths]);report.schema=JSON.parse(validation.stdout);check(validation.exit===0&&report.schema.status==='passed'&&report.schema.results.length===3&&report.schema.results.every((r:any)=>Array.isArray(r.errors)&&r.errors.length===0&&!r.exception&&!r.truncated),'Authored sample schema validation failed');
 const broken=await OpcPackage.open(await Bun.file(paths[1]!).bytes());broken.set('ppt/viewProps.xml',broken.text('ppt/viewProps.xml').replace(/<p:normalViewPr>[\s\S]*?<\/p:normalViewPr>/,'<p:normalViewPr/>'));const badPath=join(sandbox,'negative.pptx');await Bun.write(badPath,broken.toBytes());const negative=await command(['dotnet',validator,badPath]);report.negativeControl=JSON.parse(negative.stdout);check(negative.exit===1&&report.negativeControl.results[0]?.errors?.some((e:any)=>e.part==='/ppt/viewProps.xml'&&e.Id==='Sch_IncompleteContentExpectingComplex'),'Negative control did not catch missing view metadata');
 const expected=[{pages:1,text:['Native Word oracle','A second paragraph.','Answer','42','BOLDMARK','TOGGLEMARK','PLAINMARK','BOTHMARK']},{pages:2,text:['Native slide oracle','Second slide oracle','Positioned oracle box','Answer','42']},{pages:1,text:['21','22','43']}];
 for(const [i,path] of paths.entries()){
  const dir=join(sandbox,`pdf-${i}`);await mkdir(dir);await ok(['libreoffice',`-env:UserInstallation=file://${sandbox}/profile-${i}`,'--headless','--convert-to','pdf','--outdir',dir,path]);
  const name=path.split('/').at(-1)!.replace(/\.[^.]+$/,'.pdf'),pdf=join(dir,name),bytes=await Bun.file(pdf).bytes();check(bytes.length>100&&new TextDecoder().decode(bytes.slice(0,5))==='%PDF-','Missing PDF output');
  await Bun.write(join(out,name),bytes);
  const info=await ok(['pdfinfo',pdf]),pageCount=Number(info.stdout.match(/^Pages:\s+(\d+)/m)?.[1]);check(pageCount===expected[i]!.pages,'Unexpected PDF page count for '+name);
  const size=info.stdout.match(/^Page size:\s+([\d.]+) x ([\d.]+) pts/m),pageSizePoints=size?[Number(size[1]),Number(size[2])]:null;
  if(i===0)check(pageSizePoints?.[0]===792&&pageSizePoints?.[1]===612,'Landscape DOCX PDF page size mismatch');
  const text=(await ok(['pdftotext','-layout',pdf,'-'])).stdout;await Bun.write(join(out,name+'.txt'),text);
  for(const marker of expected[i]!.text)check(i===2?text.split(/\s+/).includes(marker):text.includes(marker),'Missing rendered marker '+marker+' in '+name);if(i===2)check(!text.split(/\s+/).includes('999'),'Stale formula cache rendered');
  if(i===1){const pages=text.split('\f');for(const marker of ['Native slide oracle','First page'])check(pages[1]?.includes(marker),'Reordered title-slide marker on wrong page');for(const marker of ['Second slide oracle','Positioned oracle box','Answer','42'])check(pages[0]?.includes(marker),'Reordered box/table slide marker on wrong page');}
  if(i===0){
   const markup=(await ok(['pdftohtml','-xml','-stdout','-i',pdf])).stdout;await Bun.write(join(out,'word-formatting.xml'),markup);
   const parsed=parseXml(markup.replace(/<!DOCTYPE[^>]*>/,'')),texts=elements(parsed,'text');
   const observed=formattingExpected.map(wanted=>{const nodes=texts.filter(n=>n.text===wanted.text);check(nodes.length===1,'Missing/ambiguous formatting marker');const node=nodes[0]!;const descendants=(name:string)=>{const visit=(n:any):boolean=>n.localName===name||n.children.some(visit);return visit(node);};return {text:wanted.text,bold:descendants('b'),italic:descendants('i')};});
   for(const [index,wanted]of formattingExpected.entries()){const actual=observed[index]!;check(actual.italic===wanted.italic&&(wanted.text==='TOGGLEMARK'||actual.bold===wanted.bold),'Unexpected independent font-style mismatch');}
   report.effectiveFormatting={expected:formattingExpected,observed,producer:report.versions.libreoffice,extraction:'Poppler pdftohtml bold/italic tags',matching:formattingExpected.filter((v,i)=>JSON.stringify(v)===JSON.stringify(observed[i])).length,total:4,knownDifference:observed[1]!.bold?'LibreOffice keeps bold across two paragraph-style true toggles; OOXML toggle evaluation returns false':null,scope:'Four Latin markers; no Microsoft Word or general font/script/layout equivalence'};
  }
  report.rendering.push({file:name,sha256:sha(bytes),pages:pageCount,pageSizePoints,expectedText:expected[i]!.text,textSha256:sha(new TextEncoder().encode(text)),status:'passed',visualComparison:'not performed'});
 }
 const roundtrip=join(sandbox,'roundtrip');await mkdir(roundtrip);await ok(['libreoffice',`-env:UserInstallation=file://${sandbox}/profile-calc`,'--headless','--convert-to','xlsx:Calc MS Excel 2007 XML','--outdir',roundtrip,paths[2]!]);
 const recalcPath=join(roundtrip,'book.xlsx'),recalc=await Workbook.open(recalcPath),answer=recalc.worksheet('Sheet1').getCell('A3');check(answer?.kind==='formula'&&answer.formula==='SUM(A1:A2)'&&answer.cached===43,'External formula roundtrip did not calculate43');
 const recalcBytes=await Bun.file(recalcPath).bytes();await Bun.write(join(out,'recalculated.xlsx'),recalcBytes);report.calculation={producer:report.versions.libreoffice,formula:'SUM(A1:A2)',inputs:[21,22],originalCache:999,inputCache:null,expected:43,observed:answer?.cached,outputSha256:sha(recalcBytes),status:'passed',bunCalculation:false};
 for(const [i,path]of paths.entries())check(sha(await Bun.file(path).bytes())===report.inputs[i].sha256,'Oracle changed original input');
 report.status='passed';console.log('Independent schema3/3, negative control, PDF page/text3/3 and arithmetic roundtrip passed');
}catch(error){report.status='failed';report.error=String(error);throw error;}
finally{await Bun.write(join(out,'report.json'),JSON.stringify(report,null,2)+'\n');await rm(sandbox,{recursive:true,force:true});}
