import { expect, test } from 'bun:test';
import { join } from 'node:path';
import { fixturesRoot } from '../../scripts/fixture-inputs.ts';
import { executeAcceptance, selectSharedScenarios, type StepBinding } from '../../scripts/gherkin.ts';
import { bindings } from '../acceptance/steps.ts';
const id = '@id-pptx-bun-open-save-noop';
async function run(active: StepBinding[] = bindings, change: (s: string) => string = s => s) {
  const path = 'workflows/pptx/preservation.feature';
  const feature = selectSharedScenarios(path, change(await Bun.file(join(fixturesRoot(), path)).text()), [id]);
  const count = (n: number) => ({ implemented: n, planned: 0, total: n });
  return executeAcceptance({ root: '.', features: [feature], counts: { features: count(1), scenarios: count(1), cases: count(1), steps: count(6) } }, active, 'pptx-custody-unit');
}
function refused(r: Awaited<ReturnType<typeof run>>) {
  expect(r.counts.cases.failed).toBe(1); expect(r.counts.steps.failed).toBe(1);
  expect(r.counts.steps.undefined).toBe(0); expect(r.counts.steps.ambiguous).toBe(0);
}
test('canonical Presentation path/byte opening and path save preserve the real archive', async () => {
  const result = await run(); expect(result.failures).toEqual([]); expect(result.counts.cases.passed).toBe(1); expect(result.counts.cases.planned).toBe(0);
});
test('wrong first paragraph expectation fails the bound PPTX custody assertion', async () => {
  refused(await run(bindings, s => s.replace('paragraph on its first slide is Frankenstein', 'paragraph on its first slide is Wrong')));
});

test('PPTX custody predicates reject changed byte-opened, saved and reopened archives', async () => {
  type State = import('../acceptance/pptx-custody.ts').CustodyState;
  const mutations: [string, (s: State) => void][] = [
    ['Bun Presentation opens the fixture path and separately opens its archive bytes', s => { const p = s.fromBytes!.package; p.set('ppt/slides/slide1.xml', p.text('ppt/slides/slide1.xml').replace('Frankenstein', 'Wrong')); }],
    ['the path-opened presentation is saved without edits to a new PPTX path', s => { s.saved![0] = s.saved![0]! ^ 1; }],
    ['the path-opened presentation is saved without edits to a new PPTX path', s => { s.reopened![0] = s.reopened![0]! ^ 1; }],
  ];
  for (const [step, mutate] of mutations) {
    const corrupted = bindings.map(b => b.pattern.test(step) ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => { await b.run(c, ...captures); mutate(c.state as State); } } : b);
    refused(await run(corrupted));
  }
});

test('PPTX custody temporary destination is removed on both successful and failed saves', async () => {
  const { lstat } = await import('node:fs/promises');
  type State = import('../acceptance/pptx-custody.ts').CustodyState;
  for (const failure of [false, true]) {
    let captured: State | undefined;
    const controlled = bindings.map(b => b.pattern.test('the path-opened presentation is saved without edits to a new PPTX path') ? { ...b, run: async (c: Record<string, unknown>, ...captures: string[]) => {
      captured = c.state as State;
      if (failure) captured.fromPath!.save = async () => { throw new Error('injected save failure'); };
      await b.run(c, ...captures);
    } } : b);
    const result = await run(controlled);
    if (failure) refused(result); else expect(result.failures).toEqual([]);
    expect(captured?.destination).toBeDefined();
    await expect(lstat(captured!.destination!)).rejects.toMatchObject({ code: 'ENOENT' });
    await expect(lstat(join(captured!.destination!, '..'))).rejects.toMatchObject({ code: 'ENOENT' });
  }
});
