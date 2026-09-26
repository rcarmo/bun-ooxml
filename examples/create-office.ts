/** Native authoring smoke: no fixture input and no external producer process. */
import {Document,Presentation,Workbook} from '../src/index.ts';
import {OpcPackage} from '../src/opc/package.ts';
import {mkdir} from 'node:fs/promises';
import {join} from 'node:path';
const out=join(import.meta.dir,'../artifacts/examples');await mkdir(out,{recursive:true});
const document=Document.create();document.addParagraph('Bun authored this document.',{bold:true});document.addParagraph('A second paragraph.');
const docPath=join(out,'created.docx');await document.save(docPath);
if((await Document.open(docPath)).paragraphs[0]?.text!=='Bun authored this document.')throw Error('DOCX authoring lost');
const deck=Presentation.create();deck.addTextSlide('Bun authored this slide','Native title/subtitle');deck.addTextSlide('Second slide');
const pptPath=join(out,'created.pptx');await deck.save(pptPath);
if((await Presentation.open(pptPath)).slides.length!==2)throw Error('PPTX authoring lost');
const workbook=Workbook.create();workbook.worksheet('Sheet1').setCellValue('A1','Bun authored this cell');const results=workbook.addWorksheet('Results');results.setCellValue('B3',42);
const xlsPath=join(out,'created.xlsx');await workbook.save(xlsPath);
if((await Workbook.open(xlsPath)).worksheet('Results').getCell('B3')?.value!==42)throw Error('XLSX authoring lost');
for(const path of [docPath,pptPath,xlsPath])await OpcPackage.open(path);
console.log('Native DOCX/PPTX/XLSX creation saved and reopened with expected content and package references.');
