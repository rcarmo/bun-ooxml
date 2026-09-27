/** @script Independent development-only schema, PDF text and calculation smoke.
 * @usage bun scripts/office-oracles.ts
 * @description Requires dotnet10, LibreOffice24.2.7 and Poppler; never imported by src.
 */
import {mkdir,mkdtemp,rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join,resolve} from 'node:path';
import {Document,Presentation,Workbook,OpcPackage,resolveRevisions,inspectRevisions} from '../src/index.ts';
import {parseXml,elements} from '../src/xml/index.ts';
import {runOracleCommand} from './oracle-process.ts';
import {authorStyleSample,verifyStyleReadback,styleFaults,corruptStyleSample,verifyStyleRefusal} from './xlsx-style-oracle.ts';

const root=resolve(import.meta.dir,'..'),out=join(root,'artifacts/office-oracles');
const sandbox=await mkdtemp(join(tmpdir(),'bun-office-oracles-'));
const sha=(bytes:Uint8Array)=>new Bun.CryptoHasher('sha256').update(bytes).digest('hex');
const report:any={schemaVersion:1,status:'running',sources:{},runtime:'Bun-only; external programs used here solely as test oracles',versions:{},commands:[],inputs:[],schema:null,negativeControl:null,rendering:[],calculation:null,limits:['Three rendering samples and one introduced-style workbook, not all fixtures or format operations','PDF page/text assertions do not establish visual fidelity or font equivalence','LibreOffice is not Microsoft Office','One arithmetic formula recalculated externally; no Bun calculation credit']};
await mkdir(out,{recursive:true});
async function command(args:string[],seconds=90){
 const result=await runOracleCommand(args,root,seconds*1000);report.commands.push(result);
 if(result.timedOut||result.outputLimited)throw Error('Oracle timeout/output limit: '+args[0]);return result;
}
const check=(condition:unknown,message:string)=>{if(!condition)throw Error(message);};
async function ok(args:string[],seconds=90){const r=await command(args,seconds);check(r.exit===0,`${args[0]} failed: ${r.stderr}`);return r;}
try{
 for(const source of ['src/docx/revisions.ts','src/docx/revision-properties.ts','scripts/office-oracles.ts','scripts/oracle-process.ts','tests/oracles/schema/Program.cs','tests/oracles/schema/SpreadsheetStyleReader.cs','scripts/xlsx-style-oracle.ts','src/workflow/index.ts','src/xlsx/styles.ts','tests/oracles/schema/SchemaCheck.csproj','tests/oracles/schema/packages.lock.json','src/docx/index.ts','src/opc/core-properties.ts','src/docx/style-authoring.ts','src/docx/tracking-settings.ts','src/docx/page-layout.ts','src/docx/document-properties.ts','src/docx/effective-formatting.ts','src/docx/row-header.ts','src/docx/paragraph-properties.ts','src/docx/paragraph-text.ts','src/docx/append-run.ts','src/docx/body-insertion.ts','src/docx/table-rows.ts','src/docx/table-merge.ts','src/docx/table-style.ts','src/pptx/index.ts','src/pptx/text-box.ts','src/pptx/slide-order.ts','src/xlsx/index.ts','src/xlsx/cell-style.ts'])report.sources[source]=sha(await Bun.file(join(root,source)).bytes());
 report.versions.dotnet=(await ok(['dotnet','--version'])).stdout.trim();
 report.versions.libreoffice=(await ok(['libreoffice','--version'])).stdout.trim();
 const poppler=await ok(['pdftotext','-v']);report.versions.poppler=(poppler.stderr||poppler.stdout).trim();
 const project='tests/oracles/schema/SchemaCheck.csproj';
 await ok(['dotnet','restore',project,'--locked-mode']);await ok(['dotnet','build',project,'--no-restore','--configuration','Release']);
 const validator=join(root,'tests/oracles/schema/bin/Release/net10.0/SchemaCheck.dll');
 const coreExpected={title:'Doc Title',creator:'Doc Author',subject:'Doc Subject',description:'Doc Description',keywords:'one;two',category:'Category',language:'en-US',contentStatus:'Draft',identifier:'urn:example:doc',lastModifiedBy:'Reviewer',revision:'2',version:'1.0',created:'2026-02-03T00:00:00Z',modified:'2026-02-03T01:00:00Z',lastPrinted:'2026-02-03T02:00:00Z'};
 const doc=Document.create();doc.setCoreProperties(coreExpected);doc.setDocumentProperties({titlePage:true,backgroundColor:'EEEEEE'});doc.addParagraphStyle('Smoke',{name:'Oracle heading',bold:true});doc.addParagraph('Native Word oracle',{style:'Smoke'}).setProperties({outlineLevel:0});check(doc.paragraphs[0]!.directProperties().outlineLevel===0,'Direct outline level mismatch');doc.addParagraph('Obsolete text.').appendRun(' trailing text').setText('A second paragraph.');const headingAnchor=doc.inspectBodyAnchors('Native Word oracle')[0]!;check(headingAnchor.type==='section_heading','Direct body anchor classification mismatch');doc.insertParagraphAfter(headingAnchor,'Inserted body paragraph.');
 doc.addParagraphStyle('CascadeBase',{name:'Cascade Base',bold:true});doc.addParagraphStyle('CascadeToggle',{name:'Cascade Toggle',basedOn:'CascadeBase',bold:true,italic:true});
 doc.addParagraph('BOLDMARK',{style:'CascadeBase'});doc.addParagraph('TOGGLEMARK',{style:'CascadeToggle'});doc.addParagraph('PLAINMARK',{style:'CascadeBase'});doc.paragraphs.at(-1)!.setRunFormatting({bold:false});doc.addParagraph('BOTHMARK',{bold:true,italic:true});
 const formattingExpected=doc.paragraphs.filter(p=>p.text.endsWith('MARK')).map(p=>{const r=p.effectiveRunFormatting()[0]!;return {text:r.text,bold:r.bold.value,italic:r.italic.value};});
 check(JSON.stringify(formattingExpected)===JSON.stringify([{text:'BOLDMARK',bold:true,italic:false},{text:'TOGGLEMARK',bold:false,italic:true},{text:'PLAINMARK',bold:false,italic:false},{text:'BOTHMARK',bold:true,italic:true}]),'Native cascade probe mismatch');
 let table=doc.addTable(2,2);table.setStyle('TableGrid');check(table.styleId==='TableGrid','Direct table style reference mismatch');table.cell(0,0).text='Answer';table.cell(0,1).text='42';table.setRowHeader(0,true);table.setRowHeader(1,false);check(table.isRowHeader(0)&&!table.isRowHeader(1),'Direct header readback mismatch');table=table.appendRow();table=table.insertRow(1);table=table.deleteRow(1);check(table.rows===3&&table.cell(0,1).text==='42','Row edit smoke mismatch');doc.setPageLayout({...doc.getPageLayout(),width:15840,height:12240,orientation:'landscape'});
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
 // Track preference persistence is separate from the explicit redline writer.
 const tracking=Document.create();tracking.addParagraph('Tracking settings oracle');tracking.enableTracking('Session reviewer');
 const enabledPath=join(out,'tracking-enabled.docx'),disabledPath=join(out,'tracking-disabled.docx');await tracking.save(enabledPath);tracking.disableTracking();await tracking.save(disabledPath);
 const trackingPaths=[enabledPath,disabledPath],trackingHashes=await Promise.all(trackingPaths.map(async p=>sha(await Bun.file(p).bytes())));
 const trackingValidation=await command(['dotnet',validator,...trackingPaths]);const trackingSchema=JSON.parse(trackingValidation.stdout);
 check(trackingValidation.exit===0&&trackingSchema.status==='passed'&&trackingSchema.results.length===2&&trackingSchema.results.every((r:any)=>!r.exception&&!r.truncated&&r.errors?.length===0),'Tracking settings schema validation failed');
 for(const [i,path]of trackingPaths.entries()){const readback=await Document.open(path);check(readback.trackChangesEnabled===(i===0)&&readback.trackAuthor==='','Tracking preference or session-only author readback mismatch');check(sha(await Bun.file(path).bytes())===trackingHashes[i],'Tracking validator changed input');}
 report.trackingSettings={schema:trackingSchema,sha256:trackingHashes,status:'passed',scope:'Enabled and disabled saved preferences; native reopen verifies flag and empty session author; no automatic revision or Office UI behaviour claim'};
 const mergeDoc=Document.create();mergeDoc.addTable(2,4);mergeDoc.tables[0]!.cell(0,0).text='Merged first cell';mergeDoc.tables[0]!.cell(1,3).text='Retained last cell';mergeDoc.tables[0]!.mergeRowCells(0,0,2);
 const mergePath=join(out,'merged-table.docx');await mergeDoc.save(mergePath);const mergeHash=sha(await Bun.file(mergePath).bytes());
 const mergeValidation=await command(['dotnet',validator,mergePath]),mergeSchema=JSON.parse(mergeValidation.stdout);
 check(mergeValidation.exit===0&&mergeSchema.status==='passed'&&mergeSchema.results.length===1&&mergeSchema.results[0].errors?.length===0&&!mergeSchema.results[0].exception&&!mergeSchema.results[0].truncated,'Horizontal merge schema validation failed');
 const mergeRead=await Document.open(mergePath);check(mergeRead.tables[0]!.columns===4&&mergeRead.tables[0]!.cell(1,3).text==='Retained last cell'&&mergeRead.paragraphs.some(p=>p.text==='Merged first cell'),'Horizontal merge native readback mismatch');check(sha(await Bun.file(mergePath).bytes())===mergeHash,'Merge validator changed input');
 report.horizontalMerge={schema:mergeSchema,sha256:mergeHash,status:'passed',scope:'One native horizontal merge with empty absorbed cells; schema validation and native readback, no rendering or vertical-merge claim'};
 const verticalDoc=Document.create();verticalDoc.addTable(3,3);verticalDoc.tables[0]!.cell(0,1).text='Vertical owner';verticalDoc.tables[0]!.cell(0,1).setProperties({shading:'ABCDEF',verticalAlign:'center'});verticalDoc.tables[0]!.cell(2,2).text='Vertical tail';verticalDoc.tables[0]!.mergeColumnCells(1,0,2);
 const verticalPath=join(out,'vertical-merge.docx');await verticalDoc.save(verticalPath);const verticalHash=sha(await Bun.file(verticalPath).bytes());
 const verticalValidation=await command(['dotnet',validator,verticalPath]),verticalSchema=JSON.parse(verticalValidation.stdout);
 check(verticalValidation.exit===0&&verticalSchema.status==='passed'&&verticalSchema.results.length===1&&verticalSchema.results[0].errors?.length===0&&!verticalSchema.results[0].exception&&!verticalSchema.results[0].truncated,'Vertical merge schema validation failed');
 const verticalRead=await Document.open(verticalPath);check(verticalRead.tables[0]!.rows===3&&verticalRead.tables[0]!.columns===3&&verticalRead.tables[0]!.cell(2,2).text==='Vertical tail'&&verticalRead.paragraphs.some(p=>p.text==='Vertical owner'),'Vertical merge native readback mismatch');check(sha(await Bun.file(verticalPath).bytes())===verticalHash,'Vertical validator changed input');
 report.verticalMerge={schema:verticalSchema,sha256:verticalHash,status:'passed',scope:'One single-column vertical merge with empty continuation cells; schema validity and native readback, no rendered-geometry claim'};
 const propertyDoc=Document.create();propertyDoc.addParagraph('Property revision oracle');const propertyPackage=await OpcPackage.open(propertyDoc.package.toBytes()),propertyPart=propertyPackage.mainPart();
 const propertySnapshot='<w:rPr><w:b/><w:color w:val="112233"/><w:rPrChange w:id="80" w:author="Reviewer" w:date="2026-09-27T00:00:00Z"><w:rPr><w:i/><w:color w:val="445566"/></w:rPr></w:rPrChange></w:rPr>';
 propertyPackage.set(propertyPart,propertyPackage.text(propertyPart).replace('<w:r>','<w:r>'+propertySnapshot));
 const propertySource=propertyPackage.toBytes(),propertyPaths:string[]=[];
 for(const action of ['source','accept','reject'] as const){const pkg=await OpcPackage.open(propertySource);if(action!=='source'){check(resolveRevisions(pkg,action,{profile:'text-and-run-properties'}).resolved===1,'Property resolution count mismatch');check(inspectRevisions(pkg,{profile:'text-and-run-properties'}).revisions.length===0,'Property resolution left revision');const expected=action==='accept'?'<w:rPr><w:b/><w:color w:val="112233"/></w:rPr>':'<w:rPr><w:i/><w:color w:val="445566"/></w:rPr>';check(pkg.text(propertyPart)===propertyPackage.text(propertyPart).replace(propertySnapshot,expected),'Property output differs outside selected snapshot');}const path=join(out,`run-properties-${action}.docx`);await pkg.save(path);propertyPaths.push(path);}
 const propertyHashes=await Promise.all(propertyPaths.map(async p=>sha(await Bun.file(p).bytes()))),propertyValidation=await command(['dotnet',validator,...propertyPaths]),propertySchema=JSON.parse(propertyValidation.stdout);
 check(propertyValidation.exit===0&&propertySchema.status==='passed'&&propertySchema.results.length===3&&propertySchema.results.every((r:any)=>!r.exception&&!r.truncated&&r.errors?.length===0),'Run-property revision schema validation failed');
 for(const [i,path]of propertyPaths.entries())check(sha(await Bun.file(path).bytes())===propertyHashes[i],'Property validator changed input');
 const invalidProperty=await OpcPackage.open(propertySource);invalidProperty.set(propertyPart,invalidProperty.text(propertyPart).replace(' w:author="Reviewer"',''));const invalidPropertyPath=join(sandbox,'run-properties-invalid.docx');await invalidProperty.save(invalidPropertyPath);const invalidPropertyValidation=await command(['dotnet',validator,invalidPropertyPath]),invalidPropertySchema=JSON.parse(invalidPropertyValidation.stdout);check(invalidPropertyValidation.exit===1&&invalidPropertySchema.results.some((r:any)=>r.errors?.some((e:any)=>String(e.Description??e.description??'').includes('author'))),'Missing-author property negative control did not fail');
 report.runPropertyRevisions={status:'passed',schema:propertySchema,negative:invalidPropertySchema,sha256:propertyHashes,scope:'One body run snapshot and both resolved outputs; schema validation only, no Word UI or multistory producer equivalence'};
 const sdkCore=report.schema.results.find((r:any)=>r.file==='word.docx')?.coreProperties;check(sdkCore&&Object.entries(coreExpected).every(([key,value])=>sdkCore[key]===value),'SDK core-property readback mismatch');report.coreProperties={expected:coreExpected,observed:sdkCore,status:'passed',scope:'Fifteen supplied metadata fields read by the SDK; no automatic Office author/timestamp update claim'};
 const styleSample=await authorStyleSample(out),styleBytes=await Bun.file(styleSample.output).bytes();
 const styleValidation=await command(['dotnet',validator,styleSample.output]);check(styleValidation.exit===0,'Independent style reader failed');
 const styleEvidence=JSON.parse(styleValidation.stdout),observedStyles=verifyStyleReadback(styleEvidence,styleSample.output,styleSample.sha256);
 report.styleReadback={status:'running',reader:{name:styleEvidence.validator,version:styleEvidence.version,profile:styleEvidence.profile},writer:{name:'bun-ooxml',version:JSON.parse(await Bun.file(join(root,'package.json')).text()).version,runtime:Bun.version},sourceSha256:styleSample.sourceSha256,outputSha256:styleSample.sha256,observed:observedStyles,controls:[],scope:'Saved explicit cell XFs and their referenced dependency indices, not effective formatting or rendering; optional oracle only, no default Gherkin execution credit'};
 for(const fault of styleFaults){
  const bad=await corruptStyleSample(styleBytes,fault),path=join(sandbox,`style-negative-${fault.name}.xlsx`);await Bun.write(path,bad);
  const result=await command(['dotnet',validator,path]),evidence=JSON.parse(result.stdout);check(result.exit===1,'Corrupt style reader unexpectedly succeeded');verifyStyleRefusal(evidence,path,sha(bad),fault.error);check(sha(await Bun.file(path).bytes())===sha(bad),'Style refusal changed input');
  report.styleReadback.controls.push({name:fault.name,sha256:sha(bad),status:'passed',readerResult:evidence});
 }
 check(sha(await Bun.file(styleSample.source).bytes())===styleSample.sourceSha256&&sha(await Bun.file(styleSample.output).bytes())===styleSample.sha256,'Style reader changed saved inputs');report.styleReadback.status='passed';
 const broken=await OpcPackage.open(await Bun.file(paths[1]!).bytes());broken.set('ppt/viewProps.xml',broken.text('ppt/viewProps.xml').replace(/<p:normalViewPr>[\s\S]*?<\/p:normalViewPr>/,'<p:normalViewPr/>'));const badPath=join(sandbox,'negative.pptx');await Bun.write(badPath,broken.toBytes());const negative=await command(['dotnet',validator,badPath]);report.negativeControl=JSON.parse(negative.stdout);check(negative.exit===1&&report.negativeControl.results[0]?.errors?.some((e:any)=>e.part==='/ppt/viewProps.xml'&&e.Id==='Sch_IncompleteContentExpectingComplex'),'Negative control did not catch missing view metadata');
 const expected=[{pages:1,text:['Native Word oracle','Inserted body paragraph.','A second paragraph.','Answer','42','BOLDMARK','TOGGLEMARK','PLAINMARK','BOTHMARK']},{pages:2,text:['Native slide oracle','Second slide oracle','Positioned oracle box','Answer','42']},{pages:1,text:['21','22','43']}];
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
   check(text.indexOf('Native Word oracle')<text.indexOf('Inserted body paragraph.')&&text.indexOf('Inserted body paragraph.')<text.indexOf('A second paragraph.'),'Inserted body paragraph rendered out of order');
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
 report.status='passed';console.log('Independent schema3/3, tracking-settings2/2 and horizontal-merge1/1, vertical-merge1/1, run-property-revisions3/3 with missing-author refusal, introduced-style readback with six refusal controls, PDF page/text3/3 and arithmetic roundtrip passed');
}catch(error){report.status='failed';report.error=String(error);throw error;}
finally{await Bun.write(join(out,'report.json'),JSON.stringify(report,null,2)+'\n');await rm(sandbox,{recursive:true,force:true});}
