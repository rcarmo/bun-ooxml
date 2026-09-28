import { expect, test } from 'bun:test';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { fixturePath, F } from '../../scripts/fixture-inputs.ts';
import { readZip, writeZip } from '../../src/opc/zip.ts';
import { Workbook } from '../../src/xlsx/index.ts';
import { patchOffice } from '../../src/workflow/index.ts';

const chain = 'xl/chains/order.xml';
const relType = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain';
const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml';
const encode = new TextEncoder(), decode = new TextDecoder();

async function ownedChainInput() {
  const original = Uint8Array.from(await Bun.file(fixturePath(F.shared.crossSheetCache)).bytes());
  const parts = readZip(original);
  const sheet = 'xl/worksheets/sheet2.xml', xml = decode.decode(parts.get(sheet)!);
  expect(xml).toContain('<s:f>Input!A1*2</s:f><s:v>2</s:v>');
  const enriched = xml.replace('</s:row>', '<s:c r="B1"><s:f>A1+1</s:f><s:v>3</s:v></s:c><s:c r="C1"><s:f>42</s:f><s:v>42</s:v></s:c></s:row>');
  expect(enriched).not.toBe(xml);
  parts.set(sheet, encode.encode(enriched));
  const rels = 'xl/_rels/workbook.xml.rels', relXml = decode.decode(parts.get(rels)!);
  expect(relXml).not.toContain(relType);
  const nextRels = relXml.replace('</rel:Relationships>', `<rel:Relationship Id="rIdOwnedCalcChain" Type="${relType}" Target="chains/order.xml"/></rel:Relationships>`);
  expect(nextRels).not.toBe(relXml);
  parts.set(rels, encode.encode(nextRels));
  const contentTypes = '[Content_Types].xml', types = decode.decode(parts.get(contentTypes)!);
  const nextTypes = types.replace('</ct:Types>', `<ct:Override PartName="/${chain}" ContentType="${mime}"/></ct:Types>`);
  expect(nextTypes).not.toBe(types);
  parts.set(contentTypes, encode.encode(nextTypes));
  parts.set(chain, encode.encode('<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="2"/><c r="B1" i="2"/></calcChain>'));
  return writeZip(parts);
}

test('owned nonstandard chain value edit refuses before stale metadata can survive', async () => {
  const source = await ownedChainInput(), workbook = await Workbook.open(source);
  expect(workbook.readCell('Input', 'A1')).toMatchObject({ kind: 'number', value: 1 });
  expect(workbook.readCell('Calc', 'A1')).toMatchObject({ kind: 'formula', cached: 2 });
  expect(workbook.readCell('Calc', 'B1')).toMatchObject({ kind: 'formula', cached: 3 });
  expect(workbook.readCell('Calc', 'C1')).toMatchObject({ kind: 'formula', cached: 42 });
  expect(() => workbook.worksheet('Input').setCellValue('A1', 10)).toThrow(expect.objectContaining({ code: 'xlsx-calculation-chain-unsupported' }));
  expect(workbook.readCell('Input', 'A1')).toMatchObject({ kind: 'number', value: 1 });
  expect(workbook.toBytes()).toEqual(source);
});

test('owned chain workflow refusal preserves source and an existing destination', async () => {
  const root = await mkdtemp(join(tmpdir(), 'bun-owned-chain-'));
  try {
    const source = join(root, 'source.xlsx'), output = join(root, 'output.xlsx'), bytes = await ownedChainInput();
    await Bun.write(source, bytes); await Bun.write(output, 'old output');
    const result = await patchOffice({ source, output, mode: 'safe', changes: [{ target: 'Input!A1', value: 10 }] });
    expect(result.status).toBe('refused'); expect(result.committedChanges).toBe(0);
    expect(result.error?.code).toBe('xlsx-calculation-chain-unsupported');
    expect([...await Bun.file(source).bytes()]).toEqual([...bytes]);
    expect(await Bun.file(output).text()).toBe('old output');
  } finally { await rm(root, { recursive: true, force: true }); }
});
