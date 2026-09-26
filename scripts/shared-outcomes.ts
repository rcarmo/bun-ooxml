import { isDeepStrictEqual } from "node:util";
import { PickleStepType } from "@cucumber/messages";
import type { AcceptanceFeature, AcceptanceCase, StepArgumentView } from "./gherkin.ts";
import { stableCaseKey } from "./shared-contracts.ts";

/** Cross-language evidence envelope. Native runners produce records independently.
 * Validation checks identity and assertion consistency, never Office semantics.
 * A record is useful only after its bytes and assertions receive runner/reviewer checks.
 */
export type JsonValue = null | boolean | number | string | JsonValue[] | { [key: string]: JsonValue };
export type SharedOutcome = {
  version: 2;
  runId: string;
  scenarioId: string;
  caseId: string;
  stableCaseKey: string;
  featureSha256: string;
  subject: {
    project: string; commit: string; dirty: boolean; dirtyManifestSha256?: string;
    runtime: string; version: string;
    kind: "native-library" | "direct-server-call" | "mcp-client";
    transport: "none" | "inprocess" | "stdio-mcp" | "http-mcp";
  };
  inputs: { id: string; sha256: string; origin:
    { kind: "frozen"; manifestSha256: string } | { kind: "generated"; recipeSha256: string }
  }[];
  status: "passed" | "failed";
  steps: {
    text: string; argument: StepArgumentView | null;
    status: "passed" | "failed" | "skipped" | "undefined" | "ambiguous";
    assertions: { name: string; expected: JsonValue; actual: JsonValue; status: "passed" | "failed" }[];
  }[];
  artifacts: { path: string; sha256: string }[];
};

/** The expected run ID comes from the caller's fresh run, never from this record.
 * Contract source and expanded arguments must exactly match the pinned Gherkin.
 * Then/And outcome steps need substantive expected/actual assertions to pass.
 */
export function validateSharedOutcome(record: SharedOutcome, feature: AcceptanceFeature, definition: AcceptanceCase, expectedRunId: string): void {
  requireValue(record.version===2,"Unsupported outcome version");
  requireValue(feature.lifecycle==="implemented","Planned contracts cannot claim executed outcomes");
  requireValue(!!expectedRunId && record.runId===expectedRunId,"Stale outcome run");
  requireValue(record.featureSha256===feature.sourceSha256,"Stale feature hash");
  const scenario=feature.scenarios.find(s=>s.cases.some(c=>c.identityKey===definition.identityKey));
  requireValue(!!scenario && record.scenarioId===scenario.scenarioId && record.caseId===definition.caseId,"Wrong scenario/case identity");
  requireValue(record.stableCaseKey===stableCaseKey(record.scenarioId,definition.example?.values??{}),"Wrong stable example identity");
  const subject=record.subject;
  requireValue(!!subject.project && !!subject.runtime && !!subject.version && /^[0-9a-f]{40,64}$/.test(subject.commit),"Missing subject provenance");
  requireValue(typeof subject.dirty==="boolean","Missing dirty state");
  if(subject.dirty)requireHash(subject.dirtyManifestSha256,"Dirty source manifest");
  const transports: Record<SharedOutcome["subject"]["kind"], string[]>={"native-library":["none"],"direct-server-call":["inprocess"],"mcp-client":["stdio-mcp","http-mcp"]};
  requireValue(transports[subject.kind]?.includes(subject.transport),"Subject kind/transport mismatch");
  requireValue(record.inputs.length>0,"No input fixtures recorded");
  const ids=new Set<string>();
  for(const input of record.inputs) {
    requireValue(!!input.id&&!ids.has(input.id),"Missing/duplicate input identity");ids.add(input.id);
    requireHash(input.sha256,"Input bytes");
    if(input.origin.kind==="frozen")requireHash(input.origin.manifestSha256,"Frozen manifest");
    else {requireValue(input.origin.kind==="generated","Unknown input provenance");requireHash(input.origin.recipeSha256,"Generator recipe");}
  }
  requireValue(["passed","failed"].includes(record.status),"Invalid execution status");
  requireValue(record.steps.length===definition.steps.length && record.steps.length>0,"Missing or extra steps");
  let failed=false;
  for(let i=0;i<definition.steps.length;i++) {
    const expected=definition.steps[i]!, step=record.steps[i]!;
    requireValue(step.text===expected.text && isDeepStrictEqual(step.argument,expected.argument??null),"Changed step/expanded arguments");
    requireValue(["passed","failed","skipped","undefined","ambiguous"].includes(step.status),"Unknown step status");
    if(step.status!=="passed")failed=true;
    if(expected.type===PickleStepType.OUTCOME && step.status==="passed")requireValue(step.assertions.length>0,"Passed Then requires expected/actual assertions");
    for(const assertion of step.assertions) {
      requireValue(!!assertion.name,"Unnamed assertion");
      requireValue(isJson(assertion.expected) && isJson(assertion.actual),"Assertions require explicit JSON expected/actual values");
      const equal=isDeepStrictEqual(assertion.expected,assertion.actual);
      requireValue(assertion.status===(equal?"passed":"failed"),"Assertion outcome disagrees with values");
      if(!equal) {failed=true;requireValue(step.status==="failed","Failed assertion on a non-failed step");}
    }
  }
  requireValue(record.status===(failed?"failed":"passed"),"Overall result disagrees with steps");
  const paths=new Set<string>();
  for(const artifact of record.artifacts) {
    requireValue(!!artifact.path && !artifact.path.startsWith("/") && !/[\\:\u0000-\u001f]/.test(artifact.path) && artifact.path.split("/").every(p=>!!p&&p!==".."&&p!==".") && !paths.has(artifact.path),"Unsafe/duplicate artifact path");
    paths.add(artifact.path); requireHash(artifact.sha256,"Artifact");
  }
}
function requireValue(condition: unknown, message: string): asserts condition { if(!condition)throw new Error(message); }
function requireHash(value: unknown, label: string): void { requireValue(typeof value==="string"&&/^[0-9a-f]{64}$/.test(value),`${label} SHA256 missing/invalid`); }
function isJson(value: unknown, depth=0): value is JsonValue {
  if(depth>64)return false;
  if(value===null||typeof value==="string"||typeof value==="boolean")return true;
  if(typeof value==="number")return Number.isFinite(value);
  if(Array.isArray(value))return value.every(v=>isJson(v,depth+1));
  return typeof value==="object" && !!value && [Object.prototype,null].includes(Object.getPrototypeOf(value)) && Object.values(value).every(v=>isJson(v,depth+1));
}
