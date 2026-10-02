# Uniform XML and static-reference profiles

The additive public profiles implement the sealed 18 IDs / 59 cases / 279 steps at candidate `0531afe2f0bddb50879cf7f11d5f7035e3fe374f`. Default `v0.152.0` / `28e492f` and legacy APIs are unchanged. Runtime gates and local commit await frozen source review.

## Public APIs

`analyzeStaticReferences(source)`, `parseStaticRange(source)` and `remapStaticReferences(source, contextSheet, insertion)` return detached records or a new expression. Formula spans are UTF-8 bytes. Analysis accepts the 13 functions with sealed arities and enforces 1MiB source/output, 100,000 tokens, depth128 (root0) and 10,000 references. Direct ranges preserve reversed endpoints, flags and zero absent-axis coordinates. Cell coordinates use ASCII; formula whitespace outside literals is ASCII space; direct ranges and reference tokens have no outside/internal whitespace. These APIs do not evaluate formulas or edit workbooks.

`UniformXmlSnapshot.parse(source)` returns immutable issued targets with UTF-16 half-open offsets. Methods `setAttributes`, `appendChildren`, `replaceElements` and `remove` return new Unicode source strings, which callers can encode as UTF-8. Leading BOM refuses. Limits are 8Mi UTF-16 source/output units, 100,000 elements/patches and depth256. Edits preserve quote style, unchanged tokens and deterministic namespace prefixes. A target is accepted only by its issuing snapshot; snapshots/targets remain reusable after success/refusal/no-op.

`UniformApiError.category` supplies the sealed refusal. `uniformApiResult(() => operation())` returns `{ok:true,value}` or `{ok:false,category,value:null}` for typed profile refusals. Unexpected errors propagate. Legacy error codes retain their spelling and gain an optional typed cause so output resource errors stay distinct from unsafe XML without message matching.

## Scoped evidence

The step bindings dispatch every selected predicate against production results and independently frozen full records. The matrix uses literal cell/flag goldens for 288 checks inside one canonical case. Final predicates assert exact categories, input/patch custody, snapshot/target stability and held-operation replay. Legacy bindings remain active for historical feature hashes.

Focused bindings pass59 cases/279 steps. Nine native tests pass116 assertions, including exact source/token/depth/reference boundaries, XML source/node/output limits, invalid Unicode, foreign/root/overlap refusals, frozen inputs and untrusted iterators. Five isolated production faults were assertion-red then restored: formula arity, last endpoint, foreign-target category, output-node limit category and authored XML attribute. This scope establishes no renderer, Office, formula-evaluation, package-schema or blanket parity claim. Publication/default pin movement and central execution credit are held.
