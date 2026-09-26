import {Document} from '../../src/docx/index.ts';
import {Workbook} from '../../src/xlsx/index.ts';
import {OpcPackage} from '../../src/opc/package.ts';

/** Native test builders: exercise edge topology without importing another
 * implementation's fixture corpus or serialised expected result. */
export async function textboxDocument():Promise<Uint8Array>{
 const doc=Document.create();doc.addParagraph('Outside the text box.');const pkg=await OpcPackage.open(await doc.save());
 pkg.set(pkg.mainPart(),pkg.text(pkg.mainPart()).replace('</w:body>','<w:p><w:r><w:drawing><w:txbxContent><w:p><w:r><w:t>Text living inside the text box.</w:t></w:r></w:p></w:txbxContent></w:drawing></w:r></w:p></w:body>'));
 return pkg.toBytes();
}
export async function mergedNestedTableDocument():Promise<Uint8Array>{
 const doc=Document.create(),table=doc.addTable(3,3);table.cell(0,0).text='merged';table.cell(2,0).text='nested';
 const pkg=await OpcPackage.open(await doc.save());let xml=pkg.text(pkg.mainPart());
 xml=xml.replace('<w:tcPr>','<w:tcPr><w:vMerge w:val="restart"/>');
 xml=xml.replace('<w:t>nested</w:t></w:r></w:p>','<w:t>nested</w:t></w:r></w:p><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="100"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>inner</w:t></w:r></w:p></w:tc></w:tr></w:tbl>');
 pkg.set(pkg.mainPart(),xml);return pkg.toBytes();
}
export async function sharedFormulaWorkbook():Promise<Uint8Array>{
 const wb=Workbook.create();wb.addWorksheet('Calc');const pkg=await OpcPackage.open(wb.toBytes());
 const rel=pkg.relationships(pkg.mainPart()).filter(r=>r.type.endsWith('/worksheet')).at(-1)!;
 const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
 pkg.set(rel.resolved!,`<worksheet xmlns="${ns}"><sheetData><row r="2"><c r="B2"><f t="shared" si="0" ref="B2:B3">A2*2</f><v>2</v></c><c r="D2"><f t="array" ref="D2:D3">A2:A3*2</f><v>2</v></c></row><row r="3"><c r="B3"><f t="shared" si="0"/><v>4</v></c><c r="D3"><v>4</v></c></row></sheetData></worksheet>`);
 return pkg.toBytes();
}
