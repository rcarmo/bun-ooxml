import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { OoxmlError, classifyPackageFailure } from '../../src/errors.ts';

type State = { error?: unknown; output?: unknown; archive?: Uint8Array; before?: Uint8Array };
const state = (c: Record<string, unknown>) => c.state as State;
/** Shared result predicates; archive/before are optional caller-buffer custody checks. */
export const bindings: StepBinding[] = [
  {pattern:/^(?:opening refuses with reason (\S+) and no package result|reading refuses with reason (\S+) and no member payload result|writing refuses with reason (\S+) and no archive result|saving refuses with reason (\S+) before destination replacement)$/,run:(c,...reasons)=>{
    const s=state(c),reason=reasons.find(Boolean);assert(reason);assert.equal(classifyPackageFailure(s.error),reason);assert.equal(s.output,undefined);
  }},
  {pattern:/^no successful save receipt is returned$/,run:c=>{assert(state(c).error);assert.equal(state(c).output,undefined);}},

  { pattern: /^it throws an OoxmlError with code (\S+)$/, run: (c, code) => {
    const s = state(c); assert(s.error instanceof OoxmlError); assert.equal(s.error.code, code); assert.equal(s.output, undefined);
    if (s.archive) { assert(s.before); assert.deepEqual(s.archive, s.before); }
  } },
  { pattern: /^the error message contains (.+)$/, run: (c, message) => { const s = state(c); assert(s.error instanceof OoxmlError); assert(s.error.message.includes(message!)); } },
];
