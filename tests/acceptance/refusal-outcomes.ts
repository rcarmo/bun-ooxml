import assert from 'node:assert/strict';
import type { StepBinding } from '../../scripts/gherkin.ts';
import { OoxmlError } from '../../src/errors.ts';

type State = { error?: unknown; output?: unknown; archive?: Uint8Array; before?: Uint8Array };
const state = (c: Record<string, unknown>) => c.state as State;
/** Shared result predicates; archive/before are optional caller-buffer custody checks. */
export const bindings: StepBinding[] = [
  { pattern: /^it throws an OoxmlError with code (\S+)$/, run: (c, code) => {
    const s = state(c); assert(s.error instanceof OoxmlError); assert.equal(s.error.code, code); assert.equal(s.output, undefined);
    if (s.archive) { assert(s.before); assert.deepEqual(s.archive, s.before); }
  } },
  { pattern: /^the error message contains (.+)$/, run: (c, message) => { const s = state(c); assert(s.error instanceof OoxmlError); assert(s.error.message.includes(message!)); } },
];
