import { expect, test } from 'bun:test';
import { fixturePath, F } from '../../scripts/fixture-inputs.ts';
import { OpcPackage } from '../../src/opc/package.ts';
import { inspectRetainedVisibility } from './pptx-retained-visibility.ts';

const source = F.goSlides.hiddenSlides;
const control = F.officeSlides.hiddenSlides;
const read = async (id: string) => Uint8Array.from(await Bun.file(fixturePath(id)).bytes());

// This is a package observation. Neither archive has a PowerPoint-confirmed hidden slide.
test('retained S/P slide identities and namespace-qualified marker match their unchanged inputs', async () => {
  for (const [role, id, relationships] of [
    ['source', source, ['rId7', 'rId8', 'rId9', 'rId10']],
    ['control', control, ['rId2', 'rId3', 'rId4', 'rId5']],
  ] as const) {
    const original = await read(id);
    const result = await inspectRetainedVisibility(original, role);
    expect(result).toEqual({ slideIds: ['256', '257', '258', '259'], relationships: [...relationships], parts: [1, 2, 3, 4].map(n => `ppt/slides/slide${n}.xml`) });
    expect(await read(id)).toEqual(original);
  }
});

async function corrupt(id: string, part: string, change: (xml: string) => string) {
  const original = await read(id);
  const pkg = await OpcPackage.open(original);
  const xml = pkg.text(part), changed = change(xml);
  expect(changed).not.toBe(xml);
  pkg.set(part, changed);
  expect(await read(id)).toEqual(original);
  return pkg.toBytes();
}

test('qualified marker cannot masquerade as the unqualified CT_Slide attribute', async () => {
  const part = 'ppt/slides/slide3.xml';
  const unqualified = await corrupt(source, part, xml => xml.replace('p:show="0"', 'show="0"'));
  await expect(inspectRetainedVisibility(unqualified, 'source')).rejects.toThrow();
  const controlMarker = await corrupt(control, part, xml => xml.replace('<p:sld ', '<p:sld p:show="0" '));
  await expect(inspectRetainedVisibility(controlMarker, 'control')).rejects.toThrow();
});

test('changed slide identity, relationship target and missing part fail closed', async () => {
  const changedId = await corrupt(source, 'ppt/presentation.xml', xml => xml.replace('id="256"', 'id="600"'));
  await expect(inspectRetainedVisibility(changedId, 'source')).rejects.toThrow();
  const changedTarget = await corrupt(source, 'ppt/_rels/presentation.xml.rels', xml => xml.replace('Target="slides/slide1.xml"', 'Target="slides/slide4.xml"'));
  await expect(inspectRetainedVisibility(changedTarget, 'source')).rejects.toThrow();
  await expect(corrupt(source, 'ppt/_rels/presentation.xml.rels', xml => xml.replace('Target="slides/slide4.xml"', 'Target="slides/missing.xml"'))).rejects.toThrow('Missing target');
});
